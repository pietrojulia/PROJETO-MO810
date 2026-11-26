Requirements:
openpyxl>=3.1.2
requests>=2.31.0
beautifulsoup4>=4.12.3

01_premios:
Anliasa todas as tabelas exportadas do site do Anuário sobre Prêmios. 
Gera resultados sobre as pessoas que ganharam cada prêmio e cada ano. 

02_defesas:
Primeiro script passa pelo site do Anuário, recolhe as informações de prêmios e junta com os professores presentes nas funções de orientador, coorientador e presença nas bancas. Ele gera, para cada ano, uma planilha com todas estas informações por defesa de mestrado e doutorado, entituladas: "Classificação_{ano}_preenchida.xlsx".
Porém, algumas correções foram feitas manualmente sobre a planilha. Para mantê-las, foi gerado outro script que valida o ganho dos prêmios, sem alterar as modificações feitas manualmente. Assim, o script de validação gera as mesmas planilhas, mas com os prêmios corrigidos, entituladas: "Classificação_{ano}_premios_validados.xlsx".

python classificar_defesas_premios_validados.py --atualizar-classificacao "Classificação_2025_preenchida.xlsx" --premios-base "../01_premios/resultado/categorias_ic/classificacao_detalhada.csv" --ano 2025 --saida "Classificação_2025_premios_validados.xlsx"

03_processamento:
Utiliza os das premiações, as relações de gênero e raça, produzindo uma tabela com as informações de quais docents receberam quais prêmios. Seria uma etapa de prepare dos dados para a produção do site. 

# Escolha sua banca

## 1. Como abrir no VS Code (Windows)
1. Instale as dependências uma vez:

   ```powershell
   py -m pip install -r requirements.txt
   ```

2. Execute o site:

   ```powershell
   py site_escolha_sua_banca.py
   ```

3. Abra `http://localhost:8001` se o navegador não abrir automaticamente. Pare o servidor com **Ctrl+C**.


## 2. Organização do projeto

| Arquivo | Responsabilidade |
| --- | --- |
| `index.html` | Textos, ordem das seções, títulos e estrutura |
| `estilo.css` | Cores, tipografia, espaços e versão para celular |
| `interacao.js` | Seletores, mapa de pessoas, gráficos, troféu e simulador |
| `site_escolha_sua_banca.py` | Servidor local e criação dos dados JSON |
| `preparar_dados.py` | Reconstrói participações a partir do Excel e lê prêmios exclusivamente da base validada |
| `caminhos.py` | Localiza a raiz do repositório e as cinco fontes originais |
| `dados/participacoes.csv` | Detalhes de participação gerados automaticamente |
| `dados/reconhecimentos.csv` | Registros oficiais derivados da base validada |

## Fontes de entrada (relativas à raiz do repositório)

| Fonte | Caminho |
| --- | --- |
| Anuário | `graficos/graficos_total_defesas/Anuário 2015-2025.xlsx` |
| Prêmios | `analises/01_premios/resultado/categorias_ic/classificacao_detalhada.csv` |
| Docentes | `analises/03_processamento/resultado_analise_validada/tabela_docentes_ic_filtrada.csv` |
| Defesas | `analises/03_processamento/resultado_analise_validada/Classificação_2023_2025_unificada.xlsx` |
| Cadastro | `graficos/graficos_docents_IC_2026_site/Relação gênero e raça IC.xlsx` |


Para conferir os caminhos: `python caminhos.py`.


## 3. Atualização e validação dos dados

- Execute as etapas 01_premios e 02_defesas quando houver novas entradas, depois `python analises/03_processamento/analisar_docentes_ic.py` a partir da raiz do repositório. O site lê diretamente os resultados dessas etapas, sem cópias.
- Dentro de `site_ic`, execute `python caminhos.py` para verificar as fontes.
- Execute `python preparar_dados.py` e `python conferir_dados.py`. Os testes de regressão esperam os totais do recorte atual. Em novos anos, revise conscientemente esses totais.
- Execute `python site_escolha_sua_banca.py`. A inicialização regenera os arquivos derivados e interrompe em caso de divergência. Reinicie o servidor após atualizar os dados.

## Unidades de contagem

Bancas (2023–2025): 173 defesas, 59 docentes distintos participantes, 533 pares distintos docente–defesa, 705 funções. Um orientador que também seja titular conta duas funções, mas somente um par docente–defesa. Os gráficos agregados por perfil acompanham o ano, status ou papel.

Prêmios (2019–2025): 293 registros distintos para 44 docentes, incluindo 213 premiações/colocações, 36 menções honrosas, 16 finalistas/indicações e 28 outros reconhecimentos. O subtotal de conquistas é 249. As distribuições são globais e não acompanham os seletores individuais do troféu.

**Uso local:** a API não envia gênero ou raça individual. Entretanto, pequenas contagens agregadas e dados nominais de participação/prêmios podem permitir inferências. Revise a divulgação e a proteção de pequenos grupos antes de publicar; o pacote não é uma versão pública anonimizada. O servidor disponibiliza somente os três arquivos de interface e a API, não a pasta dados.

## Teste opcional da interface

Instale `pip install playwright` e `playwright install chromium`, depois execute `python testar_interface.py`. O teste usa o JSON real produzido pelo Python e o HTML/JS/CSS reais, com transporte de API simulado, e confere o conteúdo de 45 gráficos em Chromium. Não exige um servidor HTTP em execução.
