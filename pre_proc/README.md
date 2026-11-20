# Seleção e Pré-processamento dos Dados

## 1. No VsCode ou Google Colab:
1. Execute os blocos de código em sequência.

2. O objetivo desta etapa é organizar as informações do Anuário, Bancas de Defesas e as Relações de Gênero e Raça.

3. O resultado é uma tabela unificada com os anos de 2023 a 2025 das Bancas e informações dos docentes.

## 2. Organização do projeto

| Arquivo | Responsabilidade |
| --- | --- |
| `Anuário 2025-2025.xlsx` | Informações dos anuários de 2015 a 2025 |
| `Classificação_2023_2025_unificada.xlsx` | Informações das Bancas de Defesa de Mestrado e Doutorado unificadas de acordo com o orientador e ano |
| `Classificação_2023(4,5).xlsx` | Informações das Bancas de Defesa de Mestrado e Doutorado dos anos 2023, 2024 e 2025 |
| `Relação gênero e raça IC.xlsx` | Relação de gênero e raça obtida de acordo com o site do Instituto |
| `analysing_v1.ipynb` | Transforma os dados do anuário, classificações e relações de gênero e raça em uma tabela unificada |
| `tabela_docentes_ic.csv` | **Tabela consolidada principal**, com professores sem vínculo |
| `tabela_docentes_ic_filtrada.csv` | **Tabela consolidada principal**, com informações de participações, prêmios e identificações com professores com vínculo |

