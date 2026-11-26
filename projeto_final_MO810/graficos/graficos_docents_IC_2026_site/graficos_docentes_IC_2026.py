"""Gráficos de docentes do IC/Unicamp — levantamento de 2026.
Instalação: pip install pandas matplotlib openpyxl
Execução: python graficos_docentes_IC_2026.py
Mantenha a planilha Excel na mesma pasta do script.
"""
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

# 1. Configuração e leitura somente das colunas de análise
PASTA = Path(__file__).resolve().parent if '__file__' in globals() else Path.cwd()
ARQUIVO = PASTA / 'Relação gênero e raça IC.xlsx'
SAIDA = PASTA
SAIDA.mkdir(exist_ok=True)

plt.rcParams.update({'font.size': 11, 'axes.spines.top': False,
                     'axes.spines.right': False, 'figure.dpi': 120})

dados = pd.read_excel(ARQUIVO, usecols=lambda c: c.strip() in
                      ['Professor', 'MS', 'Gênero', 'Raça', 'Status'])
dados.columns = dados.columns.str.strip()
dados = dados.dropna(how='all').copy()
# Remove espaços extras, sem modificar as classificações da coleta.
for coluna in dados.columns:
    dados[coluna] = dados[coluna].astype('string').str.strip()
    dados[coluna] = dados[coluna].replace('', pd.NA).fillna('Sem informação')

ativos = dados[dados['Status'] == 'Ativos']
ordem_status = ['Ativos', 'Aposentados', 'Ex-professores', 'Colaboradores', 'In Memoriam']
# Inclui automaticamente outras categorias, se a base for ampliada.
ordem_status += [s for s in dados['Status'].unique() if s not in ordem_status]
ordem_ms = sorted(dados['MS'].unique())
ordem_genero = ['Homem cis', 'Mulher cis']
ordem_genero += [g for g in dados['Gênero'].unique() if g not in ordem_genero]
ordem_raca = ['Branco', 'Preto', 'Pardo']
ordem_raca += [r for r in dados['Raça'].unique() if r not in ordem_raca]

# 2. Funções de visualização

def salvar(fig, ax, titulo, nota, nome):
    ax.set_title(titulo, loc='left', fontsize=16, fontweight='bold', pad=20)
    fig.text(0.015, 0.015,
             'Fonte: levantamento manual do site do IC/Unicamp, 2026.\n' + nota,
             fontsize=9)
    fig.tight_layout(rect=(0, 0.085, 1, 1))
    fig.savefig(SAIDA / (nome + '.png'), dpi=220, bbox_inches='tight')
    plt.show()
    plt.close(fig)


def barras(dados_grafico, titulo, nota, nome, tamanho=(11, 5.5), percentual=False):
    fig, ax = plt.subplots(figsize=tamanho)
    dados_grafico.plot(kind='bar', ax=ax, rot=0, width=0.76)
    for grupo in ax.containers:
        if percentual:
            total = dados_grafico.to_numpy().sum()
            rotulos = [f'{v:.0f} ({v/total:.1%})'.replace('.', ',') if total else '0'
                       for v in grupo.datavalues]
            ax.bar_label(grupo, labels=rotulos, padding=4, fontsize=10)
        else:
            ax.bar_label(grupo, fmt='%.0f', padding=4, fontsize=10)
    ax.set_xlabel('')
    ax.set_ylabel('Quantidade de docentes')
    ax.set_ylim(0, max(1, dados_grafico.max().max() * 1.25))
    ax.yaxis.set_major_locator(MaxNLocator(integer=True))
    ax.grid(axis='y', alpha=0.20)
    ax.set_axisbelow(True)
    if len(dados_grafico.columns) > 1:
        ax.legend(loc='upper right', frameon=False)
    else:
        ax.get_legend().remove()
    salvar(fig, ax, titulo, nota, nome)


def mapa_calor(matriz, titulo, nota, nome, separacoes=()):
    # Linhas = categorias; colunas = status; células = quantidades.
    fig, ax = plt.subplots(figsize=(11, max(4.6, 2.4 + 0.34 * len(matriz))))
    imagem = ax.imshow(matriz.values, aspect='auto', vmin=0, cmap='viridis')
    ax.set_xticks(range(len(matriz.columns)), matriz.columns)
    ax.set_yticks(range(len(matriz.index)), matriz.index)
    for i in range(len(matriz.index)):
        for j in range(len(matriz.columns)):
            valor = matriz.iloc[i, j]
            r, g, b, _ = imagem.cmap(imagem.norm(valor))
            cor = 'black' if 0.299*r + 0.587*g + 0.114*b > 0.52 else 'white'
            ax.text(j, i, str(int(valor)), ha='center', va='center', color=cor)
    for separacao in separacoes:
        ax.axhline(separacao - 0.5, color='white', linewidth=3)
    ax.set_xlabel('Status')
    fig.colorbar(imagem, ax=ax, label='Quantidade de docentes', pad=0.02)
    salvar(fig, ax, titulo, nota, nome)


def comparar(coluna, ordem):
    # Ativos são um subconjunto do total, não uma categoria adicional.
    return pd.DataFrame({
        'Total cadastrado': dados[coluna].value_counts(),
        'Ativos': ativos[coluna].value_counts()
    }).reindex(ordem).fillna(0).astype(int)


def por_status(coluna, ordem):
    return pd.crosstab(dados[coluna], dados['Status']).reindex(
        index=ordem, columns=ordem_status, fill_value=0)

# 3. Status — universo completo
status = dados['Status'].value_counts().reindex(ordem_status, fill_value=0)
barras(status.to_frame('Quantidade'), 'Docentes por status | 2026',
       f'Total de registros: {len(dados)}. Percentuais sobre todos os cadastrados.',
       '01_status', percentual=True)

# 4. Magistério — total e ativos
magisterio = comparar('MS', ordem_ms)
barras(magisterio, 'Docentes por nível de magistério | 2026',
       '','02_magisterio')

# 5. Gênero — total e ativos
sexo = comparar('Gênero', ordem_genero)
barras(sexo, 'Docentes por gênero | 2026',
       '',
       '03_genero', tamanho=(8, 5.5))

# 6. Heatmap de gênero e magistério por status — grupos de gênero juntos
# Primeiro todos os MS masculinos, depois todos os MS femininos.
ordem = pd.MultiIndex.from_product([ordem_genero, ordem_ms], names=['Gênero', 'MS'])
genero_ms = pd.crosstab([dados['Gênero'], dados['MS']], dados['Status'])
genero_ms = genero_ms.reindex(index=ordem, columns=ordem_status, fill_value=0)
genero_ms.index = [f'{ms} - {g}' for g, ms in genero_ms.index]
mapa_calor(genero_ms, 'Gênero por magistério e status | 2026',
           '',
           '04_genero_magisterio_status',
           separacoes=[i * len(ordem_ms) for i in range(1, len(ordem_genero))])

# 7. Raça — total e ativos; mantém Pardo? separado de Pardo
raca = comparar('Raça', ordem_raca)
barras(raca, 'Docentes por raça | 2026',
       'Classificações registradas na coleta',
       '05_raca', tamanho=(9, 5.5))

# 8. Mapas complementares por status
raca_status = por_status('Raça', ordem_raca)
mapa_calor(raca_status, 'Raça por status | 2026',
           '', '06_raca_status')
ms_status = por_status('MS', ordem_ms)
mapa_calor(ms_status, 'Magistério por status | 2026',
           'Contagem de registros em cada nível de magistério e status.', '07_magisterio_status')

# 9. Conferência — imprime apenas resultados agregados
for titulo, tabela in [('Status', status), ('Magistério', magisterio),
                       ('Gênero', sexo), ('Raça', raca)]:
    print(f'\n{titulo}:\n{tabela.to_string()}')

conferencia = pd.DataFrame({
    'Total cadastrado': [len(dados), magisterio.iloc[:, 0].sum(), sexo.iloc[:, 0].sum(),
                         raca.iloc[:, 0].sum(), genero_ms.to_numpy().sum(),
                         raca_status.to_numpy().sum(), ms_status.to_numpy().sum()],
    'Ativos': [len(ativos), magisterio['Ativos'].sum(), sexo['Ativos'].sum(),
               raca['Ativos'].sum(), genero_ms['Ativos'].sum(),
               raca_status['Ativos'].sum(), ms_status['Ativos'].sum()]
}, index=['Base', 'Magistério', 'Gênero', 'Raça', 'Gênero x MS x status',
          'Raça x status', 'MS x status'])
print('\nConferência dos totais:\n' + conferencia.to_string())
assert (conferencia['Total cadastrado'] == len(dados)).all()
assert (conferencia['Ativos'] == len(ativos)).all()
