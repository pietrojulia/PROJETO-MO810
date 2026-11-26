"""Caminhos compartilhados do site, relativos ao repositório clonado."""
from pathlib import Path

PASTA_SITE = Path(__file__).resolve().parent


def localizar_raiz():
    for pasta in (PASTA_SITE, *PASTA_SITE.parents):
        if (pasta / 'analises').is_dir() and (pasta / 'graficos').is_dir():
            return pasta
    raise FileNotFoundError(
        'Raiz do projeto não encontrada. Coloque a pasta do site dentro do '
        'repositório que contém analises/ e graficos/.'
    )


RAIZ = localizar_raiz()
PROCESSAMENTO = RAIZ / 'analises' / '03_processamento' / 'resultado_analise_validada'
ANUARIO = RAIZ / 'graficos' / 'graficos_total_defesas' / 'Anuário 2015-2025.xlsx'
BASE_PREMIOS = RAIZ / 'analises' / '01_premios' / 'resultado' / 'categorias_ic' / 'classificacao_detalhada.csv'
TABELA_DOCENTES = PROCESSAMENTO / 'tabela_docentes_ic_filtrada.csv'
CLASSIFICACAO = PROCESSAMENTO / 'Classificação_2023_2025_unificada.xlsx'
RELACAO_GENERO_RACA = RAIZ / 'graficos' / 'graficos_docents_IC_2026_site' / 'Relação gênero e raça IC.xlsx'
# Apenas os dois CSVs derivados são escritos nesta pasta, nunca as fontes.
DADOS = PASTA_SITE / 'dados'
PARTICIPACOES = DADOS / 'participacoes.csv'
RECONHECIMENTOS = DADOS / 'reconhecimentos.csv'
FONTES = {'Anuário': ANUARIO, 'Prêmios validados': BASE_PREMIOS,
          'Tabela de docentes': TABELA_DOCENTES, 'Defesas unificadas': CLASSIFICACAO,
          'Relação de gênero e raça': RELACAO_GENERO_RACA}


def verificar_fontes(*, incluir_anuario=True):
    faltantes = [(nome, arq) for nome, arq in FONTES.items()
                 if (incluir_anuario or nome != 'Anuário') and not arq.is_file()]
    if faltantes:
        detalhes = '\n'.join(f'  - {nome}: {arq.relative_to(RAIZ)}' for nome, arq in faltantes)
        raise FileNotFoundError('Fontes não encontradas no repositório:\n' + detalhes
                                + '\nExecute as etapas 01_premios, 02_defesas e 03_processamento '
                                  'quando necessário. Confira também as planilhas de graficos/.')


if __name__ == '__main__':
    verificar_fontes()
    print(f'Raiz detectada: {RAIZ}')
    for nome, arquivo in FONTES.items():
        print(f'OK: {nome}: {arquivo.relative_to(RAIZ)}')
    print(f'CSV derivados: {DADOS}')
