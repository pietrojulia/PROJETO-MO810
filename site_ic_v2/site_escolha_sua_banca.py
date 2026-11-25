"""Servidor local do projeto Escolha sua banca.

Execute no VS Code: python site_escolha_sua_banca.py
Este arquivo não consulta APIs de terceiros. Todas as entradas estão em dados/.
"""
from __future__ import annotations
from collections import Counter, defaultdict
from functools import lru_cache
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import argparse
import csv
import json
import re
import unicodedata
import webbrowser
import pandas as pd
from preparar_dados import executar as preparar_dados, premios_oficiais

ROOT = Path(__file__).resolve().parent
DADOS = ROOT / 'dados'
CATEGORIAS = ['Orientação', 'Coorientação', 'Titular', 'Suplente']
MS = ['MS1', 'MS2', 'MS3', 'MS5', 'MS6']
RACAS = ['Amarela', 'Branca', 'Indígena', 'Preta', 'Parda', 'Não Declarado']
ALIAS = {'ruben interian kovaliova': 'ruben interian'}
SITE_IC = 'https://ic.unicamp.br/docentes/'


def texto(x):
    if x is None or pd.isna(x):return ''
    return str(x).strip()


def chave(x):
    s=unicodedata.normalize('NFKD',texto(x).casefold())
    s=''.join(ch for ch in s if not unicodedata.combining(ch))
    s=re.sub(r'\s+',' ',s).strip()
    return ALIAS.get(s,s)


def ler_csv(arquivo):
    with (DADOS/arquivo).open(encoding='utf-8-sig',newline='') as f:
        return list(csv.DictReader(f,delimiter=';'))


def inteiro(v):
    try:return 0 if pd.isna(v) else int(v)
    except (TypeError,ValueError):return 0


def carregar_anuario():
    """Fonte agregada 2015–2025; preserva anos de raça sem informação."""
    sheets={k.strip():v for k,v in pd.read_excel(DADOS/'Anuario_2015_2025.xlsx',sheet_name=None).items()}
    out={}
    for ano in range(2015,2026):
        a=sheets['Docentes Ativos'].loc[sheets['Docentes Ativos']['Ano']==ano].iloc[0]
        mag=sheets['Magistério'].loc[sheets['Magistério']['Ano']==ano]
        sx=sheets['Sexo'].loc[sheets['Sexo']['Ano']==ano]
        rac=sheets['Raça'].loc[sheets['Raça']['Ano']==ano].iloc[0]
        total=inteiro(a['Total'])
        ms={nivel:sum(inteiro(v) for col in mag.columns if col==nivel or col.startswith(nivel+'.') for v in mag[col]) for nivel in MS}
        subms={col:sum(inteiro(x) for x in mag[col]) for col in mag.columns if re.fullmatch(r'MS[1-6](?:\.\d)?',col)}
        genpor={nivel:{g:sum(inteiro(x) for x in sx[f'{nivel} - {g}']) for g in ('M','F')} for nivel in MS}
        sexos={g:sum(genpor[nivel][g] for nivel in MS) for g in ('M','F')}
        racas={r:inteiro(rac[r]) for r in RACAS}
        tem_raca=not all(pd.isna(rac[r]) for r in RACAS)
        observacoes=[]
        if sum(sexos.values()) != total:
            observacoes.append(f'Sexo tabulado: {sum(sexos.values())} registros; total ativo: {total}. A diferença não foi redistribuída.')
        if sum(ms.values()) != total:
            observacoes.append('Soma das classes MS difere do total de ativos.')
        if ano==2025:
            observacoes.append('Em 2025, o regime completo tem um docente MS3.1; na aba Sexo esse registro é M.')
        if tem_raca and sum(racas.values()) != total:
            observacoes.append('A soma das categorias de raça/cor difere do total de ativos.')
        out[str(ano)]={
          'ano':ano,'total':total,'regimes':{r:inteiro(a[f'Regime {r}']) for r in ('Parcial','Completo','Integral')},
          'ms':ms,'subms':subms,'sexo':{'M':sexos['M'],'F':sexos['F'],'porMS':genpor,'registrados':sum(sexos.values())},
          'genero':None,'raca':racas if tem_raca else None,
          'status':{'Ativos':total}, 'fonte_tipo':'Anuário estatístico',
          'fonte_titulo':'Anuário Estatístico da Unicamp (2015–2025)',
          'fonte_url':'https://www.aeplan.unicamp.br/anuarios/',
          'observacoes':observacoes,
        }
    return out


def carregar_2026():
    df=pd.read_excel(DADOS/'Relacao_genero_raca_IC.xlsx',sheet_name=0).rename(columns=lambda s:str(s).strip())
    ativos=df.loc[df['Status'].astype(str).str.strip()=='Ativos'].copy()
    if len(ativos)!=47:
        raise ValueError(f'A relação de 2026 deveria conter 47 ativos, mas contém {len(ativos)}. Revise a planilha.')
    serie_ms={m:0 for m in MS};subms=Counter()
    for val in ativos['MS']:
        sub=re.sub(r'\s+','',texto(val)).upper()
        subms[sub]+=1
        base=re.match(r'MS\d+',sub)
        if base:serie_ms[base.group()]+=1
    genero=Counter(texto(x) or 'Não informado' for x in ativos['Gênero'])
    raca=Counter(texto(x) or 'Não informado' for x in ativos['Raça'])
    if sum(serie_ms.values())!=47:raise ValueError('As classes MS de 2026 não somam 47.')
    status=Counter(texto(x) for x in df['Status'])
    return {
      'ano':2026,'total':47,
      # Regimes informado pelo levantamento do grupo, não é coluna da lista nominal.
      'regimes':{'Parcial':1,'Completo':0,'Integral':46},
      'ms':serie_ms,'subms':dict(subms),
      'sexo':None,'genero':dict(genero),
      'raca':dict(raca),
      'status':dict(status),
      'fonte_tipo':'Levantamento do projeto, consulta ao site do IC',
      'fonte_titulo':'Docentes do IC/Unicamp e relação nominal do projeto',
      'fonte_url':SITE_IC,
      'observacoes':[
         'Anuário 2015–2025 não engloba o ano 2026. O site do IC mantém uma lista de docentes ativos sem série histórica oficial.',
         'A distribuição 46 integrais e 1 parcial (MS5.1, M) foi informada pelo levantamento do projeto.',
         'Gênero e raça/cor são categorias da relação preparada pelo projeto.'
      ],
    }


def contar_defesas():
    df=pd.read_excel(DADOS/'Classificação_2023_2025_unificada.xlsx',sheet_name=0)
    serie={str(ano):{'total':0,'mestrado':0,'doutorado':0} for ano in (2023,2024,2025)}
    for _,r in df.iterrows():
        try:ano=str(int(r['Ano']))
        except (TypeError,ValueError):continue
        if ano not in serie:continue
        nivel=chave(r['Classificação'])
        if nivel not in ('mestrado','doutorado') or not texto(r['Aluno']):continue
        serie[ano]['total']+=1
        serie[ano][nivel]+=1
    return serie


@lru_cache(maxsize=1)
def montar_dados():
    # Regenera e valida os detalhes; nunca usa CSVs derivados desatualizados.
    preparar_dados()
    anos=carregar_anuario();anos['2026']=carregar_2026()
    cadastro=pd.read_excel(DADOS/'Relacao_genero_raca_IC.xlsx',sheet_name=0).rename(columns=lambda c:str(c).strip())
    info={chave(row['Nome']):{'ms':texto(row['MS']),'status':texto(row['Status']),'fonteDemografia':texto(row['Lattes'])}
          for _,row in cadastro.iterrows() if texto(row['Nome'])}
    docentes=[];por_chave={}
    for r in ler_csv('tabela_docentes_ic_filtrada.csv'):
        pid=chave(r['Professor(a)'])
        if pid in por_chave:raise ValueError('Docentes com nomes duplicados depois da padronização: '+pid)
        p={'id':pid,'nome':r['Professor(a)'],'status':r['Situação no cadastro'],'genero':r['Gênero'],
            'raca':r['Raça/cor'],'ms':info.get(pid,{}).get('ms','Não informado'),
            'origemDemografia':info.get(pid,{}).get('fonteDemografia','Não consta no cadastro nominal'),
            'funcoes':{
               'Orientação':int(r['Orientador']),'Coorientação':int(r['Co-orientador']),
               'Titular':int(r['Banca titular']),'Suplente':int(r['Banca suplente'])},
            'total':int(r['Total de funções']),'bancasDistintas':int(r['Bancas distintas']),
            'premiosNumero':int(r['Quantidade de prêmios/menções']),
            'finalistasNumero':int(r['Finalistas/indicações']),'outrosNumero':int(r['Outros reconhecimentos']),
            'vinculosInferidos':int(r['Vínculos inferidos']),'observacoes':r['Observações'],
            'participacoes':[],'premios':[]}
        docentes.append(p);por_chave[pid]=p
    alertas=[]
    for r in ler_csv('participacoes.csv'):
        p=por_chave.get(chave(r['professor']))
        if p:
            p['participacoes'].append({'ano':int(r['ano']),'nivel':r['nivel'],'funcao':r['funcao'],
                         'criterio':r['criterio'],'aluno':r['aluno'],'linhaFonte':int(r['linha_fonte'])})
    for r in ler_csv('reconhecimentos.csv'):
        p=por_chave.get(chave(r['professor']))
        if p:
            p['premios'].append({'ano':int(r['ano']) if r['ano'].strip() else None,
                              'titulo':r['titulo'],'tipo':r['tipo'],'categoria':r['categoria'],'situacao':r['situacao']})
    for p in docentes:
        p['premios'].sort(key=lambda a: (a['ano'] or 0,a['titulo']),reverse=True)
        p['participacoes'].sort(key=lambda a:(-a['ano'],a['funcao']))
        funcoes_recon=Counter(x['funcao'] for x in p['participacoes'])
        diferencas={f:{'tabela':p['funcoes'][f],'detalhe':funcoes_recon[f]}
                    for f in CATEGORIAS if funcoes_recon[f]!=p['funcoes'][f]}
        if diferencas:
            alertas.append({'professor':p['nome'],'diferencas':diferencas})
        n_premios=sum(x['tipo']=='Prêmio/menção informada' for x in p['premios'])
        if n_premios!=p['premiosNumero']:
            alertas.append({'professor':p['nome'],'premios':{'tabela':p['premiosNumero'],'detalhe':n_premios}})
    # Os dados nominais do projeto contêm classificações sensíveis sem confirmação
    # pública. Calculamos distribuições aqui no servidor e NÃO enviamos gênero/raça
    # junto ao nome do docente para o navegador.
    def agregar(getter):
        grupos={'genero':Counter(),'raca':Counter(),'status':Counter()}
        for pessoa in docentes:
            quantidade=getter(pessoa)
            if quantidade<=0:continue
            grupos['genero'][pessoa['genero'] or 'Não informado']+=quantidade
            grupos['raca'][pessoa['raca'] or 'Não informado']+=quantidade
            grupos['status'][pessoa['status'] or 'Não informado']+=quantidade
        return {k:dict(v) for k,v in grupos.items()}
    if alertas:
        raise ValueError('Divergência entre os detalhes e a tabela consolidada. Atualize o notebook validado.')
    def funcoes_ano(p, y):
        return [x for x in p['participacoes'] if y=='todos' or str(x['ano'])==y]
    def distintas_ano(p, y):
        return len({(x['ano'], x['linhaFonte']) for x in funcoes_ano(p,y)})
    grupos_bancas={}
    for y in ('todos','2023','2024','2025'):
        grupos_bancas[y]={
            'docentes':agregar(lambda p:int(bool(funcoes_ano(p,y)))),
            'participacoes':agregar(lambda p:distintas_ano(p,y)),
            'funcoes':agregar(lambda p:len(funcoes_ano(p,y))),
        }
    grupos_premios={
        'registros':agregar(lambda p:len(p['premios'])),
        'docentes':agregar(lambda p:int(bool(p['premios']))),
        'conquistas':agregar(lambda p:p['premiosNumero']),
    }
    # Compatibilidade: o frontend anterior lê genero/raca/status diretamente.
    # O frontend atual usa as medidas explicitamente nomeadas.
    for grupo in grupos_bancas.values():
        grupo.update(grupo['docentes'])
    grupos_premios.update(grupos_premios['registros'])
    situacoes=Counter(x['situacao'] for p in docentes for x in p['premios'])
    categorias=Counter(x['categoria'] for p in docentes for x in p['premios'])
    if sum(situacoes.values()) != len(premios_oficiais()):
        raise ValueError('Prêmios da API não coincidem com a fonte oficial.')
    # Retirar as classificações pessoais sensíveis do JSON servido ao cliente.
    for p in docentes:
        for campo in ('genero','raca','origemDemografia'):
            p.pop(campo,None)
    docentes.sort(key=lambda p:(-p['total'],p['nome'].casefold()))
    defesas=contar_defesas()
    return {
       'versaoInterface':'2026-11-25-v2',
       'projeto':{'autores':['Andrey M. del Santo','Julia Pietro','Pedro Arruda'],
                  'periodoBancas':'2023–2025','anoConsulta':'2026'},
       'docentes':{'anos':anos,'anosDisponiveis':list(anos),'statusCadastro2026':anos['2026']['status']},
       'professores':docentes,'defesas':defesas,
       'distribuicoes':{'bancas':grupos_bancas,'premios':grupos_premios},
       'metricas':{'pessoasCadastro':len(docentes),'participacoesTabela':sum(x['total'] for x in docentes),
                   'premiosTabela':sum(x['premiosNumero'] for x in docentes),
                   'registrosPremios':sum(situacoes.values()), 'docentesPremiados':sum(bool(p['premios']) for p in docentes),
                   'situacoesPremios':dict(situacoes), 'categoriasPremios':dict(categorias),
                   'participacoesDistintas':sum(x['bancasDistintas'] for x in docentes),
                   'docentesParticipantes':sum(bool(p['participacoes']) for p in docentes),
                   'defesas':sum(x['total'] for x in defesas.values())},
       'auditoria':{'divergencias':alertas,'quantidadeDivergencias':len(alertas),
                   'mensagem':'Detalhes conferidos contra a tabela consolidada e a base oficial de prêmios.'},
       'fontes':{'anuario':'https://www.aeplan.unicamp.br/anuarios/','ic':SITE_IC},
    }


class Handler(SimpleHTTPRequestHandler):
    def __init__(self,*args,**kwargs):super().__init__(*args,directory=str(ROOT),**kwargs)
    def end_headers(self):
        self.send_header('Cache-Control','no-store, no-cache, must-revalidate, max-age=0')
        self.send_header('Pragma','no-cache')
        super().end_headers()
    def do_GET(self):
        caminho=self.path.split('?',1)[0]
        if caminho=='/api/dados':
            try:
                payload=json.dumps(montar_dados(),ensure_ascii=False).encode('utf-8')
                self.send_response(200)
            except Exception as e:
                payload=json.dumps({'erro':str(e)},ensure_ascii=False).encode('utf-8')
                self.send_response(500)
            self.send_header('Content-Type','application/json; charset=utf-8')
            self.send_header('Content-Length',str(len(payload)))
            self.end_headers();self.wfile.write(payload);return
        if caminho not in ('/', '/index.html', '/estilo.css', '/interacao.js'):
            self.send_error(403,'Recurso não publicado');return
        self.path='/index.html' if caminho=='/' else caminho
        super().do_GET()

    def do_HEAD(self):
        caminho=self.path.split('?',1)[0]
        if caminho not in ('/', '/index.html', '/estilo.css', '/interacao.js'):
            self.send_error(403,'Recurso não publicado');return
        self.path='/index.html' if caminho=='/' else caminho
        super().do_HEAD()


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--port',type=int,default=8001)
    parser.add_argument('--sem-navegador',action='store_true')
    args=parser.parse_args()
    dados=montar_dados()
    print('Dados lidos com sucesso. Defesas:',dados['metricas']['defesas'],
          '| Docentes:',dados['metricas']['pessoasCadastro'])
    if dados['auditoria']['quantidadeDivergencias']:
        print('ATENÇÃO: divergências para conferir:',dados['auditoria']['quantidadeDivergencias'])
    server=ThreadingHTTPServer(('127.0.0.1',args.port),Handler)
    endereco=f'http://localhost:{args.port}'
    print('Versão da interface: 2026-11-25-v2')
    print('Site em',endereco,'| Ctrl+C para encerrar')
    if not args.sem_navegador:webbrowser.open(endereco)
    try:server.serve_forever()
    except KeyboardInterrupt:pass
    finally:server.server_close()

if __name__=='__main__':main()
