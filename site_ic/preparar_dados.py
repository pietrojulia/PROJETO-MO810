"""Reconstroi detalhes de participacoes e reconhecimentos a partir da base de defesas.

A tabela_docentes_ic_filtrada.csv, consolidada no projeto, continua sendo a fonte dos totais
publicados. Este script verifica se os detalhes reconstruidos batem com ela.
Execute: python preparar_dados.py
"""
from collections import Counter, defaultdict, OrderedDict
from pathlib import Path
import csv
import re
import unicodedata
import pandas as pd

DADOS = Path(__file__).resolve().parent / 'dados'

# Ajustes manuais documentados, sem alterar as planilhas originais.
ALIAS = {'ruben interian kovaliova': 'ruben interian'}
NOMES_CONJUNTOS = {'alexandre xavier falcao e jean cousty': ['Alexandre Xavier Falcão']}
FUNCOES = [(i, 'Orientação' if i == 2 else 'Coorientação' if i == 7 else
            'Titular' if i <= 32 else 'Suplente') for i in (2,7,12,17,22,27,32,37,42,47)]
SEM_PREMIO = ('nao localizado','nao verificado','n/a','nao se aplica',
              'sem registro','nenhum','nao ha','sem premio')


def texto(valor):
    return '' if pd.isna(valor) else ' '.join(str(valor).split())


def normalizar(valor):
    s = unicodedata.normalize('NFKD', texto(valor).casefold())
    return ' '.join(''.join(c for c in s if not unicodedata.combining(c)).split()).replace('’', "'")


def chave(valor):
    return ALIAS.get(normalizar(valor), normalizar(valor))


def tipo_reconhecimento(titulo):
    s = normalizar(titulo)
    if re.search(r'\b(finalist\w*|nominee|nomination|indicado\w*|indicacao|shortlist\w*|candidate|contender|entre (os|as) \d+ (melhores|finalistas)|selecionad\w* como finalista|selecionad\w* entre)\b', s):
        return 'Finalista/indicação'
    if re.search(r'\b(keynote speaker|palestrant\w* convidad\w*|chair\b|distinguished lecturer|top[ -]?\d+|world scientists|top \d+%|rank\w*|fellow\b|elected member|produtividade em pesquisa|bolsa de produtividade|escolhid\w*|chosen for|representative and influential papers|distinguished reviewer|revisor de destaque|distinguished pc member|counselor|researchers in the world|scientists by c-score|scientific impact ranking|eleicao|eleito para|elected|convidado|convidada)\b', s):
        return 'Outros reconhecimentos'
    return 'Prêmio/menção informada'


def salvar_csv(nome, cabecalho, linhas):
    with (DADOS / nome).open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=cabecalho, delimiter=';', extrasaction='ignore')
        writer.writeheader()
        writer.writerows(linhas)


def executar():
    cad = pd.read_excel(DADOS/'Relacao_genero_raca_IC.xlsx', sheet_name=0)
    cadastro = {chave(row['Nome']):texto(row['Nome']) for _,row in cad.iterrows() if texto(row['Nome'])}
    base = pd.read_excel(DADOS/'Classificação_2023_2025_unificada.xlsx', sheet_name=0)
    if len(base.columns) != 53:
        raise ValueError('Esperadas 53 colunas na planilha de defesas. Confira a ordem das colunas.')
    defesas = [(idx+2, row.tolist()) for idx,row in base.iterrows()
               if normalizar(row.iloc[0]) in ('mestrado','doutorado') and texto(row.iloc[1])]
    explicit = defaultdict(Counter)
    for linha_excel, row in defesas:
        for i,_ in FUNCOES:
            nome, inst=texto(row[i]),normalizar(row[i+1]).replace(' ','')
            if nome and inst == 'ic/unicamp':
                explicit[chave(nome)][nome]+=1
    rows, pendencias=[],[]
    for linha_excel,row in defesas:
        ano=int(row[52]);nivel=texto(row[0]);aluno=texto(row[1])
        for i,funcao in FUNCOES:
            informado=texto(row[i]);inst=texto(row[i+1]);pr=texto(row[i+2]);
            if not informado or normalizar(informado) in ('nao ha','n/a','nenhum','sem suplentes','-'):
                continue
            nomes=NOMES_CONJUNTOS.get(normalizar(informado),[informado])
            for nome in nomes:
                cod=chave(nome); vinculo=normalizar(inst).replace(' ','')
                if vinculo == 'ic/unicamp': criterio='IC explícito'
                elif not vinculo and cod in cadastro: criterio='IC inferido pelo cadastro'
                elif not vinculo and cod in explicit: criterio='IC inferido por outra defesa'
                elif vinculo in ('unicamp','universidadeestadualdecampinas') and cod in cadastro:
                    criterio='Unicamp sem unidade; IC pelo cadastro'
                else:
                    if not inst: pendencias.append({'ano':ano,'aluno':aluno,'nome':nome,'funcao':funcao,'motivo':'Vínculo não comprovado'})
                    continue
                nome_padrao=cadastro.get(cod) or (explicit[cod].most_common(1)[0][0] if explicit[cod] else nome)
                rows.append({'professor':nome_padrao,'chave':cod,'ano':ano,'nivel':nivel,'aluno':aluno,'funcao':funcao,
                             'criterio':criterio,'linha_fonte':linha_excel,'premio_bruto':pr})
    if len(rows) != len({(r['chave'],r['linha_fonte'],r['funcao']) for r in rows}):
        raise ValueError('Participações duplicadas: revisar nomes e funções.')
    premios=OrderedDict()
    for r in rows:
        bruto=r['premio_bruto']
        if not bruto or normalizar(bruto).startswith(SEM_PREMIO):continue
        for parte in re.split(r'\s*\|\s*',bruto):
            parte=texto(parte)
            if not parte or normalizar(parte).startswith(SEM_PREMIO):continue
            m=re.match(r'^((?:19|20)\d{2})\s*[—–\-:]\s*(.*)$',parte)
            ano=int(m.group(1)) if m else None
            titulo=texto(m.group(2) if m else parte)
            if not titulo:continue
            key=(r['chave'],ano,normalizar(titulo))
            if key not in premios:
                premios[key]={'professor':r['professor'],'chave':r['chave'],'ano':ano or '',
                              'titulo':titulo,'tipo':tipo_reconhecimento(titulo)}
    salvar_csv('participacoes.csv',['professor','chave','ano','nivel','aluno','funcao','criterio','linha_fonte'],rows)
    salvar_csv('reconhecimentos.csv',['professor','chave','ano','titulo','tipo'],premios.values())
    with (DADOS/'tabela_docentes_ic_filtrada.csv').open(encoding='utf-8-sig',newline='') as f:
        resumo=list(csv.DictReader(f,delimiter=';'))
    contagens=defaultdict(Counter)
    for r in rows:contagens[r['chave']][r['funcao']]+=1
    erros=[]
    colunas={'Orientação':'Orientador','Coorientação':'Co-orientador','Titular':'Banca titular','Suplente':'Banca suplente'}
    for r in resumo:
        c=contagens[chave(r['Professor(a)'])]
        for funcao,col in colunas.items():
            if c[funcao] != int(r[col]):
                erros.append(f"{r['Professor(a)']}: {funcao}: detalhes={c[funcao]}, tabela={r[col]}")
    print(f'Defesas válidas: {len(defesas)} | Participações: {len(rows)} | Reconhecimentos: {len(premios)}')
    print('Prêmios/menções classificados:',sum(v['tipo']=='Prêmio/menção informada' for v in premios.values()))
    if erros:
        print(f'ATENÇÃO: {len(erros)} divergências entre detalhes e tabela consolidada.')
        for erro in erros[:12]: print('  ',erro)
    else:
        print('Conferência OK: contagens por função coincidem com o CSV consolidado.')
    print(f'Vínculos não confirmados e não incluídos: {len(pendencias)}')

if __name__=='__main__': executar()
