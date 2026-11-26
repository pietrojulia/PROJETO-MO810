"""Gráficos do Anuário de docentes (2015–2025).
Instalação: pip install pandas matplotlib openpyxl
Execução: python graficos_docentes_atualizado.py
Mantenha o arquivo Excel na mesma pasta deste script.
"""
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

# 1. Configuração e leitura da planilha
PASTA = Path(__file__).resolve().parent if '__file__' in globals() else Path.cwd()
ARQUIVO = PASTA / 'Anuário 2015-2025.xlsx'
# Aceita também o nome do arquivo original enviado.
if not ARQUIVO.exists():
    ARQUIVO = PASTA / 'Anuário 2015-2025(6).xlsx'
SAIDA = PASTA
SAIDA.mkdir(exist_ok=True)

plt.rcParams.update({'font.size': 11, 'axes.spines.top': False,
                     'axes.spines.right': False, 'figure.dpi': 120})


def ler_aba(nome):
    # A aba Docentes Ativos possui um espaço no final do nome.
    with pd.ExcelFile(ARQUIVO) as excel:
        aba = next(s for s in excel.sheet_names if s.strip() == nome)
        dados = pd.read_excel(excel, sheet_name=aba)
    # Exclui notas, linhas em branco e outros registros fora do período.
    dados['Ano'] = pd.to_numeric(dados['Ano'], errors='coerce')
    dados = dados[dados['Ano'].between(2015, 2025)].copy()
    dados['Ano'] = dados['Ano'].astype(int)
    return dados


# 2. Funções de visualização reutilizáveis

def salvar(fig, ax, titulo, nota, nome):
    ax.set_title(titulo, loc='left', fontsize=17, fontweight='bold', pad=22)
    fig.text(0.01, 0.015, 'Fonte: Anuário 2015–2025 | ' + nota, fontsize=9)
    fig.tight_layout(rect=(0, 0.055, 1, 1))
    fig.savefig(SAIDA / (nome + '.png'), dpi=220, bbox_inches='tight')
    plt.show()
    plt.close(fig)


def barras(dados, titulo, nota, nome, tamanho=(15, 6)):
    fig, ax = plt.subplots(figsize=tamanho)
    dados.plot(kind='bar', ax=ax, width=0.83, rot=0)
    # Exibe a quantidade de cada barra, inclusive zeros.
    for grupo in ax.containers:
        ax.bar_label(grupo, fmt='%.0f', padding=3, fontsize=9)
    ax.set_xlabel('Ano')
    ax.set_ylabel('Quantidade de docentes')
    ax.set_ylim(0, max(1, dados.max().max() * 1.20))
    ax.yaxis.set_major_locator(MaxNLocator(integer=True))
    ax.grid(axis='y', alpha=0.20)
    ax.set_axisbelow(True)
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, 1.01),
              ncol=len(dados.columns), frameon=False, fontsize=10)
    salvar(fig, ax, titulo, nota, nome)


def mapa_calor(dados, titulo, nota, nome, separacao=None):
    # Linhas = categorias; colunas = anos; células = quantidades.
    fig, ax = plt.subplots(figsize=(13, 6))
    matriz = dados.T
    imagem = ax.imshow(matriz.values, aspect='auto', vmin=0, cmap='viridis')
    ax.set_xticks(range(len(matriz.columns)), matriz.columns)
    ax.set_yticks(range(len(matriz.index)), matriz.index)
    for i in range(len(matriz.index)):
        for j in range(len(matriz.columns)):
            valor = matriz.iloc[i, j]
            # Contraste automático para o número ficar legível.
            r, g, b, _ = imagem.cmap(imagem.norm(valor))
            cor = 'black' if 0.299*r + 0.587*g + 0.114*b > 0.52 else 'white'
            ax.text(j, i, str(int(valor)), ha='center', va='center', color=cor)
    # Linha branca opcional entre os grupos de sexo.
    if separacao is not None:
        ax.axhline(separacao - 0.5, color='white', linewidth=3)
    ax.set_xlabel('Ano')
    fig.colorbar(imagem, ax=ax, label='Quantidade de docentes', pad=0.02)
    salvar(fig, ax, titulo, nota, nome)


# 3. Docentes ativos por regime e total
regime = ler_aba('Docentes Ativos').set_index('Ano').sort_index()
regime = regime[['Regime Parcial', 'Regime Completo', 'Regime Integral', 'Total']]
regime = regime.fillna(0).astype(int)
regime.columns = ['Parcial', 'Completo', 'Integral', 'Total']
barras(regime, 'Docentes ativos por regime de trabalho | 2015–2025',
       'Parcial, completo e integral', '01_regimes')

# 4. Magistério: soma dos três regimes de trabalho em cada ano
magisterio = ler_aba('Magistério')
colunas_ms = [c for c in magisterio.columns if c.startswith('MS')]
magisterio = magisterio.groupby('Ano')[colunas_ms].sum().astype(int)
mapa_calor(magisterio, 'Docentes ativos por nível de magistério | 2015–2025',
           'Soma dos regimes', '02_magisterio')

# 5. Sexo: soma das colunas terminadas em M e F
sexo = ler_aba('Sexo')
# Correção informada: MS5 feminino, regime integral, 2025.
mascara = (sexo['Ano'] == 2025) & (sexo['Regime'].str.strip() == 'Integral')
sexo.loc[mascara, 'MS5 - F'] = 2
colunas_sexo = [c for c in sexo.columns if c.endswith((' - M', ' - F'))]
sexo_ms = sexo.groupby('Ano')[colunas_sexo].sum().astype(int)
sexo_total = pd.DataFrame({
    'Masculino': sexo_ms.filter(regex=r' - M$').sum(axis=1),
    'Feminino': sexo_ms.filter(regex=r' - F$').sum(axis=1)
})
barras(sexo_total, 'Docentes ativos por sexo | 2015–2025',
       'Soma do sexo informado no Anuário | M = Masculino e F = Feminino',
       '03_sexo')

# 6. Sexo por nível de magistério: masculino primeiro, feminino depois
colunas_m = [c for c in sexo_ms.columns if c.endswith(' - M')]
colunas_f = [c for c in sexo_ms.columns if c.endswith(' - F')]
sexo_ms_agrupado = sexo_ms[colunas_m + colunas_f]

mapa_calor(sexo_ms_agrupado, 'Sexo por nível de magistério | 2015–2025',
           'M = masculino; F = feminino',
           '04_sexo_magisterio', separacao=len(colunas_m))

# 7. Raça: exclui anos totalmente sem informação ANTES de preencher zeros
raca = ler_aba('Raça').set_index('Ano').sort_index()
raca = raca[['Amarela', 'Branca', 'Indígena', 'Preta', 'Parda', 'Não Declarado']]
raca = raca.dropna(how='all').fillna(0).astype(int)
barras(raca, 'Docentes ativos por raça | 2023–2025',
       'Ausência de dados entre 2015 e 2022.',
       '05_raca', tamanho=(12, 6))

# 8. Conferência dos totais (após a correção informada)
conferencia = pd.DataFrame({'Docentes ativos': regime['Total'],
                           'Soma dos regimes': regime[['Parcial','Completo','Integral']].sum(axis=1),
                           'Magistério': magisterio.sum(axis=1),
                           'Sexo': sexo_total.sum(axis=1),
                           'Raça': raca.sum(axis=1)})
print('\nConferência dos totais anuais (NaN = sem informação):')
print(conferencia.to_string())
