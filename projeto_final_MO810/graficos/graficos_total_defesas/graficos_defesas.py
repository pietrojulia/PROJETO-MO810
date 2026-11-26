"""Defesas do IC/Unicamp | Anuário 1974–2025.
Instalação: pip install pandas matplotlib openpyxl
Execução: python graficos_defesas.py
Coloque a planilha na mesma pasta do script.
"""
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

PASTA = Path(__file__).resolve().parent if '__file__' in globals() else Path.cwd()
ARQUIVO = PASTA / 'Anuário 2015-2025.xlsx'
SAIDA = PASTA
SAIDA.mkdir(exist_ok=True)
plt.rcParams.update({'font.size': 11, 'axes.spines.top': False,
                     'axes.spines.right': False, 'figure.dpi': 120})

# 1. Leitura: mantém somente as linhas com anos ou períodos históricos.
modalidades = ['Mestrado', 'Doutorado', 'Mestrado Profissional']
base = pd.read_excel(ARQUIVO, sheet_name='Defesas')
base.columns = base.columns.str.strip()
base[modalidades] = base[modalidades].apply(pd.to_numeric, errors='coerce')
anos = pd.to_numeric(base['Ano'], errors='coerce')
mask = anos.between(1990, 2025)
defesas = base.loc[mask, modalidades].copy()
defesas.index = anos[mask].astype(int)
defesas.index.name = 'Ano'
defesas = defesas.sort_index()
# Não preencher NaN: falta de informação não é igual a zero.
periodos = ['1974 a 1979', '1980 a 1989']
historico = base.loc[base['Ano'].isin(periodos)].set_index('Ano')[modalidades]
historico = historico.reindex(periodos)

# 2. Função para finalizar os gráficos.
def salvar(fig, ax, titulo, nota, nome):
    ax.set_title(titulo, loc='left', fontsize=17, fontweight='bold', pad=18)
    fig.text(0.01, 0.018, 'Fonte: Anuário do IC/Unicamp | ' + nota, fontsize=9)
    fig.tight_layout(rect=(0, 0.07, 1, 1))
    fig.savefig(SAIDA / (nome + '.png'), dpi=220, bbox_inches='tight')
    plt.show()
    plt.close(fig)

# 3. Evolução anual: linhas, marcadores e destaque dos maiores valores.
fig, ax = plt.subplots(figsize=(16, 6))
for modalidade in modalidades:
    serie = defesas[modalidade]
    linha, = ax.plot(serie.index, serie, marker='o', markersize=4,
                    linewidth=2, label=modalidade)
    if serie.notna().any():
        ano_pico = serie.idxmax()
        valor_pico = serie.max()
        ax.annotate(f'{int(valor_pico)} ({ano_pico})', (ano_pico, valor_pico),
                    xytext=(0, -17 if modalidade == 'Mestrado Profissional' else 10),
                    textcoords='offset points', ha='center',
                    color=linha.get_color(), fontweight='bold', fontsize=10)
ax.set_xticks(defesas.index)
ax.tick_params(axis='x', labelrotation=60)
ax.set_xlabel('Ano')
ax.set_ylabel('Quantidade de defesas')
ax.set_ylim(0, defesas.max().max() * 1.24)
ax.yaxis.set_major_locator(MaxNLocator(integer=True))
ax.grid(axis='y', alpha=0.2)
ax.legend(loc='upper left', frameon=False, ncol=3)
salvar(fig, ax, 'Evolução anual das defesas | 1990–2025',
       'Mestrado profissional: sem informação de 2008 a 2025.',
       '01_evolucao_defesas')

# 4. Heatmap: todas as quantidades anuais, com ausências em cinza.
fig, ax = plt.subplots(figsize=(21, 4.5))
matriz = defesas.T
cmap = plt.get_cmap('viridis').copy()
cmap.set_bad('#d9d9d9')
imagem = ax.imshow(matriz.to_numpy(dtype=float), aspect='auto', vmin=0, cmap=cmap)
ax.set_xticks(range(len(matriz.columns)), matriz.columns, rotation=60)
ax.set_yticks(range(len(matriz.index)), matriz.index)
for i in range(len(matriz.index)):
    for j in range(len(matriz.columns)):
        valor = matriz.iloc[i, j]
        if pd.isna(valor):
            texto, cor = '–', 'black'
        else:
            r, g, b, _ = imagem.cmap(imagem.norm(valor))
            cor = 'black' if 0.299*r + 0.587*g + 0.114*b > 0.52 else 'white'
            texto = str(int(valor))
        ax.text(j, i, texto, ha='center', va='center', color=cor, fontsize=10)
ax.set_xlabel('Ano')
fig.colorbar(imagem, ax=ax, label='Quantidade de defesas', pad=0.02)
salvar(fig, ax, 'Defesas por modalidade e ano | 1990–2025',
       'Cinza (–) = sem informação; zero = valor registrado.', '02_heatmap_defesas')

# 5. Totais históricos: inclui os dois períodos agregados sem duplicar
# as linhas de totais já existentes na planilha.
totais = pd.concat([historico, defesas]).sum(min_count=1)
fig, ax = plt.subplots(figsize=(10, 5))
barras = ax.bar(totais.index, totais.values)
ax.bar_label(barras, fmt='%.0f', padding=4, fontsize=13, fontweight='bold')
ax.set_ylabel('Quantidade de defesas registradas')
ax.set_ylim(0, totais.max() * 1.18)
ax.grid(axis='y', alpha=0.2)
ax.set_axisbelow(True)
salvar(fig, ax, 'Total histórico de defesas registradas | 1974–2025',
       'Inclui 1974–1979 (6) e 1980–1989 (36), ambos de mestrado.',
       '03_totais_historicos')

# 6. Conferência e exportação dos valores anuais.
# Total registrado é parcial quando alguma modalidade não foi informada.
resumo = defesas.copy()
resumo['Total registrado'] = defesas.sum(axis=1, min_count=1)
resumo['Modalidades sem informação'] = defesas.isna().sum(axis=1)
print('\nDefesas anuais (NaN = sem informação):')
print(resumo.to_string())
print('\nPeríodos históricos agregados:')
print(historico.to_string())
print('\nTotais históricos registrados:')
print(totais.to_string())
print(f'Total geral registrado: {int(totais.sum())}')
# Confere com os totais informados na aba, quando encontrados.
linhas_total = base.loc[base['Ano'].isna(), modalidades].dropna(how='any')
if not linhas_total.empty:
    informado = linhas_total.iloc[0]
    print('\nDiferença em relação aos totais da planilha:')
    print((totais - informado).to_string())
resumo.to_csv(SAIDA / 'defesas_anuais.csv', sep=';', encoding='utf-8-sig')
