#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Pré-processamento das defesas e premiações do IC/UNICAMP.
Dependências: python3 -m pip install pandas openpyxl
"""
import argparse
import sys
import tempfile
from collections import Counter, defaultdict, OrderedDict
from pathlib import Path
import csv
import unicodedata
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

def localizar_raiz():
    """Encontra a raiz do repositório a partir da localização do script."""
    inicio = Path(__file__).resolve().parent
    for pasta in (inicio, *inicio.parents):
        if (pasta / 'analises' / '02_defesas').is_dir() and (pasta / 'graficos').is_dir():
            return pasta
    raise FileNotFoundError(
        'Não foi possível localizar a raiz do projeto a partir de '
        f'{inicio}. Mantenha a estrutura analises/ e graficos/ do repositório '
        'ou informe --raiz /caminho/do/projeto.'
    )


def argumentos(argv=None):
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--raiz', type=Path, default=None, help='Raiz do projeto (padrão: detectada a partir do script)')
    p.add_argument('--defesas', type=Path, help='Pasta das classificações anuais')
    p.add_argument('--relacao', type=Path, help='Planilha de cadastro nominal')
    p.add_argument('--premios', type=Path, help='CSV classificacao_detalhada.csv validado')
    p.add_argument('--saida', type=Path, default=Path(__file__).resolve().parent / 'resultado_analise_validada', help='Pasta de saída; padrão: ao lado do script')
    p.add_argument('--anos', nargs='+', type=int, default=[2023, 2024, 2025], help='Anos das defesas (não restringe os anos dos prêmios)')
    return p.parse_args(argv)


def conferir(condicao, mensagem):
    if not condicao:
        raise ValueError(mensagem)


def ler_defesas(pasta, anos):
    lista = []
    for ano in anos:
        arquivo = pasta / f'Classificação_{ano}_premios_validados.xlsx'
        if not arquivo.is_file():
            raise FileNotFoundError(f'Classificação anual não encontrada: {arquivo}')
        df = pd.read_excel(arquivo, sheet_name=0)
        colunas = list(df.columns)
        esperado = ['Classificação', 'Aluno']
        for i, papel in enumerate(['Orientador', 'Co-orientador'] + [f'Titular {j}' for j in range(1, 6)] + [f'Suplente {j}' for j in range(1, 4)]):
            sufixo = f'.{i}' if i else ''
            esperado.extend([papel, 'Instituto'+sufixo, 'Prêmio'+sufixo, 'Gênero'+sufixo, 'Raça'+sufixo])
        conferir(len(colunas) == 52 and set(colunas) == set(esperado), f'{arquivo}: esperadas as 52 colunas do modelo original.')
        conferir(colunas[:2] == esperado[:2], f'{arquivo}: cabeçalhos iniciais fora de ordem.')
        for i in range(2, 52, 5):
            conferir(colunas[i:i+3] == esperado[i:i+3] and set(colunas[i+3:i+5]) == set(esperado[i+3:i+5]), f'{arquivo}: bloco de membro fora do formato esperado, coluna {i+1}.')
        df['Ano'] = ano
        lista.append(df)
        print(f'Defesas {ano}: {len(df)} linhas — {arquivo}')
    # Concatenação por cabeçalho, não por posição: mantém a ordem do primeiro ano.
    return pd.concat(lista, ignore_index=True)


def executar(args):
    raiz = args.raiz.expanduser().resolve() if args.raiz is not None else localizar_raiz()
    pasta_defesas = (args.defesas or raiz / 'analises' / '02_defesas').expanduser().resolve()
    relacao = (args.relacao or raiz / 'graficos' / 'graficos_docents_IC_2026_site' / 'Relação gênero e raça IC.xlsx').expanduser().resolve()
    BASE_PREMIOS = (args.premios or raiz / 'analises' / '01_premios' / 'resultado' / 'categorias_ic' / 'classificacao_detalhada.csv').expanduser().resolve()
    destino = args.saida.expanduser().resolve()
    anos = sorted(set(args.anos))
    periodo = str(anos[0]) if len(anos)==1 else f'{anos[0]}–{anos[-1]}'
    nome_unificada = f'Classificação_{anos[0]}_{anos[-1]}_unificada.xlsx' if len(anos)>1 else f'Classificação_{anos[0]}_unificada.xlsx'
    for arquivo in (relacao, BASE_PREMIOS):
        if not arquivo.is_file():
            raise FileNotFoundError(f'Arquivo de entrada não encontrado: {arquivo}')
    print('Relação:', relacao)
    print('Base de prêmios:', BASE_PREMIOS)
    print('Saída:', destino)
    df_unificado = ler_defesas(pasta_defesas, anos)
    df_gen = pd.read_excel(relacao, sheet_name=0)
    conferir([str(c).strip() for c in df_gen.columns[:7]] == ['Nome','Professor','MS','Gênero','Raça','Lattes','Status'], 'Cadastro nominal: confira os sete primeiros cabeçalhos.')
    print('Cadastro nominal carregado:', len(df_gen), 'linhas')
    # Escreve em área temporária; só publica após todas as verificações.
    with tempfile.TemporaryDirectory(prefix='pre_proc_ic_') as tmp:
        PASTA_SAIDA = Path(tmp)
        def saida(nome):
            return PASTA_SAIDA / nome
        df_unificado.to_excel(saida(nome_unificada), index=False)
        # Funções de padronização de grafias e leituras de linhas

        # Os números de linha começam em 2, porque o cabeçalho ocupa a primeira linha do Excel.
        def converter_linhas(df):
            linhas = []
            for indice, valores in enumerate(df.itertuples(index=False, name=None), start=2):
                valores = [None if pd.isna(v) else v for v in valores]
                linhas.append((indice, valores))
            return linhas

        def limpar(valor):
            return ' '.join(str(valor).split()) if valor is not None else ''

        def normalizar(valor):
            sem_acento = unicodedata.normalize('NFKD', limpar(valor).casefold())
            return ' '.join(''.join(c for c in sem_acento if not unicodedata.combining(c)).split()).replace('’', "'")

        def celula(linha, i):
            return linha[i] if i < len(linha) else None

        # Correções manuais de grafias equivalentes
        ALIAS = {'ruben interian kovaliova': 'ruben interian'}
        def chave(nome):
            t = normalizar(nome)
            return ALIAS.get(t, t)

        cadastro = {}
        for linha_excel, campos in converter_linhas(df_gen):
            nome = limpar(celula(campos, 0))
            if not nome:
                continue
            conferir(chave(nome) not in cadastro, f"Nome duplicado no cadastro: {nome}")
            cadastro[chave(nome)] = {
                'nome': nome,
                'cargo': limpar(celula(campos, 1)),
                'ms': limpar(celula(campos, 2)),
                'genero': limpar(celula(campos, 3)),
                'raca_cor': limpar(celula(campos, 4)),
                'origem_genero_raca': limpar(celula(campos, 5)),
                'situacao': limpar(celula(campos, 6)),
                'linha_cadastro': linha_excel,
            }
        print('Nomes na relação de gênero/raça:', len(cadastro))



        # Identificar defesas e funções dos docentes
        defesas_xlsx = [(1, df_unificado.columns.tolist())] + converter_linhas(df_unificado)
        cabecalho = defesas_xlsx[0][1]
        assert len(cabecalho) == 53, 'Confira se a estrutura da classificação unificada mudou.'

        FUNCOES = [(str(cabecalho[i]), i, i+1, i+2) for i in (2,7,12,17,22,27,32,37,42,47)]

        defesa_validas = []
        linhas_descartadas = []
        for linha_excel, linha in defesas_xlsx[1:]:
            nivel = normalizar(celula(linha, 0))
            estudante = limpar(celula(linha, 1))
            if nivel not in ('mestrado', 'doutorado') or not estudante:
                linhas_descartadas.append(linha_excel)
                continue
            ano = int(celula(linha, 52))
            defesa_validas.append((linha_excel, linha, ano))

        # Nome explicitamente associado a IC/UNICAMP em qualquer defesa.
        nomes_ic_explicitos = Counter()
        grafias_ic = defaultdict(Counter)
        for linha_excel, linha, ano in defesa_validas:
            for funcao, i_nome, i_inst, i_premio in FUNCOES:
                nome = limpar(celula(linha, i_nome))
                instituto = normalizar(celula(linha, i_inst)).replace(' ', '')
                if nome and instituto == 'ic/unicamp':
                    nomes_ic_explicitos[chave(nome)] += 1
                    grafias_ic[chave(nome)][nome] += 1

        print('Defesas válidas:', len(defesa_validas))
        print('Funções no arquivo:', [x[0] for x in FUNCOES])

        # Pegando as participações dos docentes com vínculo com IC
        # Caso peculiar da planilha: uma célula contém duas pessoas.
        # A primeira é conhecida do IC, e a segunda requer análise separada.
        NOMES_CONJUNTOS = {
            normalizar('Alexandre Xavier Falcão e Jean Cousty'): ['Alexandre Xavier Falcão'],
            normalizar('Leonardo Montecchi e Julio Cesar dos Reis'): ['Leonardo Montecchi', 'Julio Cesar dos Reis'],
        }

        def verificar_vinculo(nome, instituto):
            cod = chave(nome)
            uni = normalizar(instituto).replace(' ', '')
            if uni == 'ic/unicamp':
                return True, 'IC explícito'
            if not uni and cod in cadastro:
                return True, 'IC inferido pelo cadastro'
            if not uni and cod in nomes_ic_explicitos:
                return True, 'IC inferido por outra defesa'
            if uni in ('unicamp', 'universidadeestadualdecampinas') and cod in cadastro:
                return True, 'UNICAMP sem unidade; IC pelo cadastro'
            return False, 'Vínculo IC não comprovado'

        participacoes = []
        pendencias = []
        for linha_excel, linha, ano in defesa_validas:
            nivel = limpar(celula(linha, 0))
            aluno = limpar(celula(linha, 1))
            for funcao, i_nome, i_inst, i_premio in FUNCOES:
                nome_original = limpar(celula(linha, i_nome))
                instituto = limpar(celula(linha, i_inst))
                premios_brutos = limpar(celula(linha, i_premio))
                if not nome_original or normalizar(nome_original) in ('nao ha','n/a','nenhum','sem suplentes','-'):
                    continue
                partes = NOMES_CONJUNTOS.get(normalizar(nome_original), [nome_original])
                if partes != [nome_original]:
                    pendencias.append([ano, aluno, nome_original, funcao, instituto, linha_excel,
                                        'Nome conjunto: analisados os nomes do IC separadamente; célula original preservada'])
                for nome in partes:
                    incluir, justificativa = verificar_vinculo(nome, instituto)
                    if not incluir:
                        if not instituto:
                            pendencias.append([ano, aluno, nome, funcao, instituto, linha_excel,
                                                'Instituto vazio sem confirmação de vínculo IC'])
                        continue
                    cod = chave(nome)
                    nome_padrao = (cadastro[cod]['nome'] if cod in cadastro
                                   else grafias_ic[cod].most_common(1)[0][0] if grafias_ic[cod]
                                   else nome)
                    participacoes.append({
                        'chave': cod, 'nome': nome_padrao, 'ano': ano, 'nivel': nivel,
                        'aluno': aluno, 'funcao': funcao, 'instituto': instituto,
                        'criterio_vinculo': justificativa, 'linha_excel': linha_excel,
                        'premios_brutos': premios_brutos,
                    })
                    if 'inferido' in justificativa or 'sem unidade' in justificativa:
                        pendencias.append([ano, aluno, nome, funcao, instituto, linha_excel,
                                            'Incluído com vínculo inferido: conferir'])

        # Uma participação = uma combinação de docente + defesa + função.
        ids = [(p['chave'],p['linha_excel'],p['funcao']) for p in participacoes]
        assert len(ids) == len(set(ids)), 'Há funções repetidas para a mesma pessoa/defesa.'
        print('Participações IC:',len(participacoes))

        # Registro único de prêmios — fonte oficial validada, não as células das bancas.
        # Não usa regex para classificar prêmios nem exige presença em uma banca.
        if not BASE_PREMIOS.is_file():
            raise FileNotFoundError(f'Base oficial não encontrada: {BASE_PREMIOS}. Use --premios para indicar o caminho correto.')
        base_premios = pd.read_csv(BASE_PREMIOS, sep=';', encoding='utf-8-sig', dtype=str, keep_default_na=False)
        obrigatorias = {'Ano', 'Professor / nome', 'Prêmio', 'Categoria', 'Situação da premiação'}
        if not obrigatorias.issubset(base_premios.columns):
            raise ValueError(f'Colunas faltantes na base: {obrigatorias - set(base_premios.columns)}. Use classificacao_detalhada.csv.')
        MAPA_SITUACAO = {
            'Premiação / colocação': 'Prêmio/menção informada',
            'Menção honrosa': 'Prêmio/menção informada',
            'Finalista / indicação': 'Finalista/indicação',
            'Outro reconhecimento': 'Outros reconhecimentos',
        }
        linhas_por_prof = defaultdict(set)
        for p in participacoes:
            linhas_por_prof[p['chave']].add(p['linha_excel'])
        premios_unicos = OrderedDict()
        for r in base_premios.to_dict('records'):
            nome = limpar(r['Professor / nome'])
            cod = chave(nome)
            ano_premio = int(r['Ano'])
            titulo = r['Prêmio']
            situacao = r['Situação da premiação']
            categoria = r['Categoria']
            if not cod or not titulo.strip() or not categoria.strip() or situacao not in MAPA_SITUACAO:
                raise ValueError(f'Registro oficial incompleto ou situação desconhecida: {nome}, {ano_premio}, {situacao}')
            id_premio = (cod, ano_premio, normalizar(titulo))
            if id_premio in premios_unicos:
                raise ValueError(f'Chave duplicada na base oficial: {id_premio}. Confira a fonte, sem descartar automaticamente.')
            premios_unicos[id_premio] = {
                'chave': cod, 'nome': nome, 'ano': ano_premio, 'titulo': titulo,
                'categoria': categoria, 'situacao': situacao, 'tipo': MAPA_SITUACAO[situacao],
                # Linhas de participação do docente, não fontes originais do prêmio.
                'linhas_fonte': linhas_por_prof[cod].copy(),
            }
        # Impede que o filtro final descarte um docente premiado silenciosamente.
        sem_cadastro = sorted({p['nome'] for p in premios_unicos.values()
                              if p['chave'] not in cadastro or cadastro[p['chave']]['situacao'] == 'Verificar vínculo'})
        if sem_cadastro:
            raise ValueError(f'Docentes da base oficial ausentes ou pendentes no cadastro: {sem_cadastro}. Revise o cadastro.')
        assert len(premios_unicos) == len(base_premios)
        print('Registros oficiais de prêmios:', len(premios_unicos))
        print('Docentes premiados:', len({p['chave'] for p in premios_unicos.values()}))
        print('Categorias originais:', base_premios['Categoria'].nunique())
        print(base_premios['Situação da premiação'].value_counts().to_string())
        # Exporta todos os campos originais para rastreabilidade, sem reclassificar.
        base_premios.to_csv(saida('premios_docentes_detalhados.csv'), sep=';', index=False, encoding='utf-8-sig')
        resumo_premios_categoria = pd.crosstab(base_premios['Categoria'], base_premios['Situação da premiação'])
        resumo_premios_categoria = resumo_premios_categoria.reindex(columns=list(MAPA_SITUACAO), fill_value=0)
        resumo_premios_categoria['Total'] = resumo_premios_categoria.sum(axis=1)
        resumo_premios_categoria.to_csv(saida('resumo_premios_por_categoria.csv'), sep=';', encoding='utf-8-sig')


        # Contabilizando por docente
        pessoas = {cod: dados.copy() for cod, dados in cadastro.items()}
        for p in participacoes:
            if p['chave'] not in pessoas:
                pessoas[p['chave']] = {
                    'nome': p['nome'], 'cargo': 'Não informado', 'ms': '',
                    'genero': 'Não informado', 'raca_cor': 'Não informado',
                    'origem_genero_raca': 'Ausente da relação nominal',
                    'situacao': 'Verificar vínculo', 'linha_cadastro': '',
                }

        participacoes_por_prof = defaultdict(list)
        premios_por_prof = defaultdict(list)
        for p in participacoes:
            participacoes_por_prof[p['chave']].append(p)
        for premio in premios_unicos.values():
            premios_por_prof[premio['chave']].append(premio)

        COLUNAS = [
            'Professor(a)', 'Situação no cadastro', 'Gênero', 'Raça/cor',
            'Orientador', 'Co-orientador', 'Banca titular', 'Banca suplente',
            'Bancas distintas', 'Total de funções', 'Quantidade de prêmios/menções',
            'Quais prêmios/menções', 'Finalistas/indicações', 'Outros reconhecimentos',
            'Anos de defesas', 'Vínculos inferidos', 'Observações'
        ]

        resumo = []
        for cod, pessoa in sorted(pessoas.items(), key=lambda par: normalizar(par[1]['nome'])):
            linhas = participacoes_por_prof[cod]
            premios = premios_por_prof[cod]
            contagem = Counter(p['funcao'] for p in linhas)
            principais = [a for a in premios if a['tipo']=='Prêmio/menção informada']
            bancas_distintas = len({p['linha_excel'] for p in linhas})
            inferidos = sum('inferido' in p['criterio_vinculo'].lower()
                            or 'sem unidade' in p['criterio_vinculo'].lower() for p in linhas)
            observacoes = []
            if cod not in cadastro:
                observacoes.append('Fora do cadastro nominal de gênero e raça')
            if not linhas:
                observacoes.append(f'Sem participação identificada em {periodo}')
            if inferidos:
                observacoes.append('Há participações com vínculo IC inferido')
            resumo.append({
                'Professor(a)': pessoa['nome'],
                'Situação no cadastro': pessoa['situacao'],
                'Gênero': pessoa['genero'],
                'Raça/cor': pessoa['raca_cor'],
                'Orientador': contagem['Orientador'],
                'Co-orientador': contagem['Co-orientador'],
                'Banca titular': sum(n for funcao,n in contagem.items() if funcao.startswith('Titular')),
                'Banca suplente': sum(n for funcao,n in contagem.items() if funcao.startswith('Suplente')),
                'Bancas distintas': bancas_distintas,
                'Total de funções': len(linhas),
                'Quantidade de prêmios/menções': len(principais),
                'Quais prêmios/menções': ' | '.join((str(a['ano'])+' — ' if a['ano'] else '') + a['titulo'] for a in principais),
                'Finalistas/indicações': sum(a['tipo']=='Finalista/indicação' for a in premios),
                'Outros reconhecimentos': sum(a['tipo']=='Outros reconhecimentos' for a in premios),
                'Anos de defesas': ', '.join(str(a) for a in sorted({p['ano'] for p in linhas})),
                'Vínculos inferidos': inferidos,
                'Observações': '; '.join(observacoes),
            })

        print('Pessoas na tabela:', len(resumo))
        print('Total de orientações:', sum(p['Orientador'] for p in resumo))
        print('Total de coorientações:', sum(p['Co-orientador'] for p in resumo))
        print('Total de participações como titular:', sum(p['Banca titular'] for p in resumo))
        print('Total de participações como suplente:', sum(p['Banca suplente'] for p in resumo))
        # Salvando o resultado
        with open(saida('tabela_docentes_ic.csv'), 'w', encoding='utf-8-sig', newline='') as arq:
            escritor = csv.DictWriter(arq, fieldnames=COLUNAS, delimiter=';')
            escritor.writeheader()
            escritor.writerows(resumo)
        print('CSV salvo em: tabela_docentes_ic.csv')

        with pd.ExcelWriter(saida('tabela_docentes_ic_com_codigo.xlsx'), engine='openpyxl') as writer:
            def aba(nome, titulos, linhas, larguras=None):
                pd.DataFrame(linhas, columns=titulos).to_excel(writer, sheet_name=nome, index=False)
                ws = writer.sheets[nome]
                ws.freeze_panes = 'A2'
                ws.auto_filter.ref = ws.dimensions
                for cell in ws[1]:
                    cell.font = Font(bold=True, color='FFFFFF')
                    cell.fill = PatternFill('solid', fgColor='155E75')
                    cell.alignment = Alignment(wrap_text=True, vertical='center')
                ws.row_dimensions[1].height = 35
                for j, w in enumerate(larguras or [], 1):
                    ws.column_dimensions[get_column_letter(j)].width = w
            aba('Resumo', COLUNAS, [[linha[c] for c in COLUNAS] for linha in resumo],
                [39,32,17,18,15,17,17,18,17,18,25,80,21,23,21,20,58])

            DETALHE = ['Professor(a)','Ano','Nível','Aluno(a)','Função','Instituto','Tipo de vínculo','Linha da unificada']
            aba('Participações', DETALHE,
                [[p[c] for c in ('nome','ano','nivel','aluno','funcao','instituto','criterio_vinculo','linha_excel')]
                 for p in participacoes], [39,12,17,43,18,22,41,24])

            PREMIOS_COL = ['Professor(a)','Ano prêmio','Título','Categoria','Situação da premiação','Linhas de participação na unificada']
            aba('Prêmios', PREMIOS_COL,
                [[p['nome'],p['ano'],p['titulo'],p['categoria'],p['situacao'],', '.join(map(str,sorted(p['linhas_fonte'])))]
                 for p in premios_unicos.values()], [39,17,86,40,30,40])

            aba('Pendências', ['Ano','Aluno(a)','Nome informado','Função','Instituto','Linha','Motivo'],
                pendencias, [14,42,44,21,25,14,72])

            NOTAS = [
                ['Período das participações',f'Defesas de mestrado e doutorado {periodo} da classificação unificada'],
                ['Gênero e raça','Somente valores transcritos da relação de gênero e raça do IC; sem inferência'],
                ['Vínculo IC','Explícito quando instituto IC/UNICAMP; algumas ausências incluídas por cadastro/ocorrência anterior'],
                ['Bancas distintas','Cada combinação professor e linha de defesa contada uma vez, independentemente de múltiplos papéis'],
                ['Prêmios','Fonte única: classificacao_detalhada.csv; não depende de participação em banca'],
                ['Deduplicação dos prêmios','Base oficial já deduplicada; chave validada sem excluir registros'],
                ['Finalistas/outros reconhecimentos','Contados à parte e não somados a prêmios/menções'],
                ['Ausência de registro','Zero não significa ausência histórica de participação ou prêmio'],
                ['Ano de prêmio','Ano da base oficial; período completo, inclusive prêmios posteriores à defesa'],
                ['Validação','Conferir as abas Pendências, Participações e Prêmios antes de publicar os dados'],
            ]
            aba('Metodologia', ['Regra','Explicação'], NOTAS, [35,110])


        print('Excel salvo em: tabela_docentes_ic_com_codigo.xlsx')

        assert sum(r['Total de funções'] for r in resumo) == len(participacoes)
        assert sum(r['Quantidade de prêmios/menções'] for r in resumo) == sum(
            p['tipo'] == 'Prêmio/menção informada' for p in premios_unicos.values())
        assert len(resumo) == len(pessoas)
        print('Verificações passaram!')
        print('Resumo: pessoas=', len(resumo), 'defesas=', len(defesa_validas),
              'participações=', len(participacoes),
              'prêmios/menções=', sum(p['Quantidade de prêmios/menções'] for p in resumo))
        print('Exemplo (primeiros 3 docentes):')
        for r in resumo[:3]:
            print(r['Professor(a)'], ':', r['Orientador'], 'orient.,',
                  r['Co-orientador'], 'coorient.,', r['Quantidade de prêmios/menções'], 'prêmios')
        tabela_docentes = pd.read_csv(saida('tabela_docentes_ic.csv'), sep=";")
        tabela_docentes = tabela_docentes[
            tabela_docentes["Situação no cadastro"] != "Verificar vínculo"
        ].copy()

        # Reorganizar os índices
        tabela_docentes.reset_index(drop=True, inplace=True)

        print("Total de professores:", len(tabela_docentes))
        print(tabela_docentes["Situação no cadastro"].value_counts())

        # Validação da coerência inclusive após o filtro final, por docente e situação.
        assert tabela_docentes['Professor(a)'].map(chave).is_unique
        por_nome = tabela_docentes.set_index(tabela_docentes['Professor(a)'].map(chave))
        campos = {'Prêmio/menção informada': 'Quantidade de prêmios/menções',
                  'Finalista/indicação': 'Finalistas/indicações',
                  'Outros reconhecimentos': 'Outros reconhecimentos'}
        for cod, linha in por_nome.iterrows():
            esperado = Counter(p['tipo'] for p in premios_unicos.values() if p['chave'] == cod)
            for tipo, coluna in campos.items():
                assert int(linha[coluna]) == esperado[tipo], (cod, coluna)
        assert {p['chave'] for p in premios_unicos.values()}.issubset(por_nome.index)
        assert sum(int(tabela_docentes[c].sum()) for c in campos.values()) == len(base_premios)
        print('Coerência validada: todos os registros oficiais presentes no CSV filtrado.')
        print(tabela_docentes[list(campos.values())].sum().to_string())

        tabela_docentes.to_csv(
            saida("tabela_docentes_ic_filtrada.csv"),
            sep=";",
            index=False,
            encoding="utf-8-sig"
        )
        # Indicadores finais, sem confundir docentes, funções e prêmios.
        print('Docentes participantes:', int((tabela_docentes['Total de funções'] > 0).sum()))
        print('Participações distintas:', int(tabela_docentes['Bancas distintas'].sum()))
        print('Funções exercidas:', int(tabela_docentes['Total de funções'].sum()))
        print('Total de registros oficiais:', len(premios_unicos))
        # Nenhum arquivo de entrada é alterado; os arquivos gerados são atualizados no destino.
        destino.mkdir(parents=True, exist_ok=True)
        import shutil
        for arquivo in sorted(PASTA_SAIDA.iterdir()):
            shutil.copy2(arquivo, destino / arquivo.name)
        print('Resultados gravados em:', destino)


def main(argv=None):
    if not __debug__:
        print('Não execute com python -O: as verificações do notebook usam assert.', file=sys.stderr)
        return 2
    try:
        executar(argumentos(argv))
    except (FileNotFoundError, ValueError, AssertionError, KeyError, OSError) as erro:
        print(f'ERRO: {erro}', file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
