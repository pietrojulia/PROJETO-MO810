"""Reconstrói participações das defesas e lê reconhecimentos da base oficial validada.

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
NOMES_CONJUNTOS = {'alexandre xavier falcao e jean cousty': ['Alexandre Xavier Falcão'], 'leonardo montecchi e julio cesar dos reis': ['Leonardo Montecchi', 'Julio Cesar dos Reis']}
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


MAPA_SITUACAO = {
    'Premiação / colocação': 'Prêmio/menção informada',
    'Menção honrosa': 'Prêmio/menção informada',
    'Finalista / indicação': 'Finalista/indicação',
    'Outro reconhecimento': 'Outros reconhecimentos',
}


def premios_oficiais():
    base = pd.read_csv(DADOS/'classificacao_detalhada.csv', sep=';', encoding='utf-8-sig', dtype=str, keep_default_na=False)
    obrigatorias = {'Ano', 'Professor / nome', 'Prêmio', 'Categoria', 'Situação da premiação'}
    if not obrigatorias <= set(base.columns):
        raise ValueError(f'Colunas ausentes da base oficial: {obrigatorias - set(base.columns)}')
    premios = OrderedDict()
    for r in base.to_dict('records'):
        nome = texto(r['Professor / nome']); ano = int(r['Ano']); titulo = r['Prêmio']; situacao = r['Situação da premiação']; categoria = r['Categoria']
        key = (chave(nome), ano, normalizar(titulo))
        if key in premios or not key[0] or not titulo.strip() or not categoria.strip() or situacao not in MAPA_SITUACAO:
            raise ValueError(f'Registro oficial duplicado ou inválido: {key}')
        premios[key] = {'professor':nome,'chave':key[0],'ano':ano,'titulo':titulo,'tipo':MAPA_SITUACAO[situacao],'categoria':categoria,'situacao':situacao}
    return premios


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
    premios=premios_oficiais()
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
    ids = {chave(r['Professor(a)']) for r in resumo}
    if {r['chave'] for r in premios.values()} - ids:
        raise ValueError('Há docentes premiados ausentes do cadastro consolidado.')
    cont_premios = defaultdict(Counter)
    for r in premios.values(): cont_premios[r['chave']][r['tipo']] += 1
    for r in resumo:
        cod = chave(r['Professor(a)'])
        for tipo, col in [('Prêmio/menção informada','Quantidade de prêmios/menções'),('Finalista/indicação','Finalistas/indicações'),('Outros reconhecimentos','Outros reconhecimentos')]:
            if cont_premios[cod][tipo] != int(r[col]):
                erros.append(f"{r['Professor(a)']}: {col} diverge da base oficial")
        distintas = len({v['linha_fonte'] for v in rows if v['chave']==cod})
        if distintas != int(r['Bancas distintas']): erros.append(f"{r['Professor(a)']}: bancas distintas divergem")
    # Exporta somente participações dos docentes no cadastro filtrado.
    rows = [r for r in rows if r['chave'] in ids]
    if erros:
        raise ValueError('Divergências: atualize o CSV pelo notebook validado.\n'+'\n'.join(erros))
    salvar_csv('participacoes.csv',['professor','chave','ano','nivel','aluno','funcao','criterio','linha_fonte'],rows)
    salvar_csv('reconhecimentos.csv',['professor','chave','ano','titulo','tipo','categoria','situacao'],premios.values())
    print(f'Defesas válidas: {len(defesas)} | Participações: {len(rows)} | Reconhecimentos: {len(premios)}')
    print('Prêmios/menções classificados:',sum(v['tipo']=='Prêmio/menção informada' for v in premios.values()))
    if erros:
        print(f'ATENÇÃO: {len(erros)} divergências entre detalhes e tabela consolidada.')
        for erro in erros[:12]: print('  ',erro)
    else:
        print('Conferência OK: contagens por função coincidem com o CSV consolidado.')
    print(f'Vínculos não confirmados e não incluídos: {len(pendencias)}')

if __name__=='__main__': executar()
