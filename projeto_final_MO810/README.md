# Escolha sua banca — Feminismo de Dados

Projeto final da disciplina **MO810 — Tópicos em Inteligência Artificial: Feminismo de Dados**, desenvolvido no Instituto de Computação (IC) da Universidade Estadual de Campinas (Unicamp).

**Equipe:** Andrey M del Santo, Julia Pietro e Pedro Arruda.

O projeto investiga quem participa das bancas de mestrado e doutorado do IC e como as participações e os registros de reconhecimento acadêmico se distribuem entre os docentes, considerando gênero, raça/cor, função e vínculo institucional. O site *Escolha sua banca* apresenta os dados de forma interativa, com gráficos, filtros, fichas de docentes e um simulador de composição de bancas.

## 1. Estrutura do projeto

O diretório `projeto_final_MO810/`, dentro do repositório, contém as etapas de coleta, processamento, análise e visualização. A árvore abaixo destaca os principais arquivos do fluxo de trabalho.

```text
PROJETO-MO810/
└── projeto_final_MO810/
    ├── README.md
    ├── requirements.txt
    ├── analises/
    │   ├── 01_premios/
    │   │   ├── consolidar_premios_ic_validado.py
    │   │   ├── Tabela de Prêmios e Homenagens ... .xlsx
    │   │   └── resultado/
    │   │       ├── premios_docentes_ic.csv
    │   │       ├── resumo_anual.csv
    │   │       ├── resumo_professor_ano.csv
    │   │       └── categorias_ic/
    │   │           ├── classificacao_detalhada.csv
    │   │           └── resumo_categorias.csv
    │   ├── 02_defesas/
    │   │   ├── classificar_defesas.py
    │   │   ├── classificar_defesas_premios_validados.py
    │   │   ├── Histórico de Defesas de Mestrado e Doutorado.xlsx
    │   │   ├── Classificação.xlsx
    │   │   ├── Classificação_{ano}_preenchida.xlsx
    │   │   └── Classificação_{ano}_premios_validados.xlsx
    │   └── 03_processamento/
    │       ├── analisar_docentes_ic.py
    │       └── resultado_analise_validada/
    │           ├── Classificação_2023_2025_unificada.xlsx
    │           ├── tabela_docentes_ic.csv
    │           ├── tabela_docentes_ic_filtrada.csv
    │           ├── tabela_docentes_ic_com_codigo.xlsx
    │           ├── premios_docentes_detalhados.csv
    │           └── resumo_premios_por_categoria.csv
    ├── graficos/
    │   ├── graficos_total_defesas/
    │   │   └── Anuário 2015-2025.xlsx
    │   └── graficos_docents_IC_2026_site/
    │       └── Relação gênero e raça IC.xlsx
    └── site_ic/
        ├── caminhos.py
        ├── preparar_dados.py
        ├── conferir_dados.py
        ├── site_escolha_sua_banca.py
        ├── testar_interface.py
        ├── index.html
        ├── estilo.css
        ├── interacao.js
        ├── requirements.txt
        └── dados/                         # Gerado automaticamente
            ├── participacoes.csv
            └── reconhecimentos.csv
```

`{ano}` representa os anos 2023, 2024 e 2025. Os caminhos são relativos ao projeto, sem depender do usuário ou do diretório em que o repositório foi clonado. Os nomes de scripts acima correspondem às versões de referência utilizadas na documentação; caso algum arquivo tenha sido renomeado no repositório, ajuste o nome no comando correspondente.

## 2. Instalação e execução rápida

Recomenda-se Python 3.10 ou superior, Git e um navegador web. O VS Code é opcional. Clone o repositório e entre no diretório do projeto:

```bash
git clone https://github.com/pietrojulia/PROJETO-MO810.git
cd PROJETO-MO810/projeto_final_MO810
```


### Dependências

O `requirements.txt` da raiz do projeto deve reunir as dependências de todas as etapas:

```text
openpyxl>=3.1.2
requests>=2.31.0
beautifulsoup4>=4.12.3
pandas>=2.1
```

Instale as dependências uma vez:

```bash
python -m pip install -r requirements.txt
```


### Abrir o site com os dados já processados

Se as cinco fontes indicadas na seção 4 estiverem disponíveis, **não é necessário repetir a coleta nem a consolidação** para visualizar o site. A partir de `projeto_final_MO810/`:

```bash
python site_ic/caminhos.py
python site_ic/site_escolha_sua_banca.py
```

Abra [http://localhost:8001](http://localhost:8001) caso o navegador não abra automaticamente. O servidor prepara os arquivos derivados e verifica a consistência dos dados na inicialização.

## 3. Etapas de coleta e processamento

Execute esta seção quando for necessário reconstruir ou atualizar as bases. A ordem é: **01_premios → 02_defesas → 03_processamento → 04_site**. Os arquivos de `graficos/` fornecem informações complementares ao processamento e ao site.

### 3.1. `analises/01_premios` — Premiações e reconhecimentos

Consolida as 14 tabelas de prêmios e homenagens internos e externos, exportadas do [Anuário de Pesquisa da Unicamp](https://pesquisa.dados.unicamp.br/premios), referentes a 2019–2025.

O script `consolidar_premios_ic_validado.py` lê as exportações Excel, identifica registros do IC, normaliza nomes, elimina duplicidades e aplica o critério de vínculo docente validado pelo projeto. São incluídos os registros a partir do ano de ingresso do docente, inclusive; registros anteriores ao ingresso e nomes ausentes da relação docente adotada para essa etapa são excluídos. Também classifica os registros por categoria e situação e sinaliza descrições novas para revisão manual.

A partir de `projeto_final_MO810/`:

```bash
cd analises/01_premios
python consolidar_premios_ic_validado.py --entrada . --saida resultado
cd ../..
```

O parâmetro `--entrada` deve apontar para a pasta que contém as exportações originais, caso estejam em um subdiretório. Não é necessário acessar o site do Anuário durante essa execução se as exportações já estiverem disponíveis.

Principais resultados:

| Arquivo em `resultado/` | Conteúdo |
| --- | --- |
| `premios_docentes_ic.csv` | Registros dos docentes após a consolidação inicial. |
| `todos_registros_ic.csv` e `outros_registros_ic.csv` | Registros consolidados e registros excluídos pelo critério docente, para auditoria. |
| `resumo_anual.csv` | Quantidades por ano e tipo (interno/externo). |
| `resumo_professor_ano.csv` | Quantidades por docente e ano. |
| `categorias_ic/classificacao_detalhada.csv` | **Base de referência validada** para as demais etapas, com pessoa, ano, título, categoria e situação de cada reconhecimento. |
| `categorias_ic/resumo_categorias.csv` | Distribuição dos registros por categoria. |

A base final possui 293 registros referentes a 44 docentes, distribuídos em 28 categorias. As etapas seguintes utilizam `classificacao_detalhada.csv`, e não os totais da consolidação inicial. Ao incorporar novos dados, revise a relação docente, as regras e as classificações sinalizadas antes de utilizar os resultados.

### 3.2. `analises/02_defesas` — Classificação das bancas

Analisa as defesas de mestrado e doutorado do IC realizadas entre 2023 e 2025. O processamento ocorre em duas etapas:

**Coleta inicial — `classificar_defesas.py`.** Lê a planilha histórica de defesas e acessa as páginas individuais de cada defesa para identificar orientadores, coorientadores, membros titulares, suplentes e suas afiliações. Também consulta as informações de prêmios do Anuário. Gera uma planilha `Classificação_{ano}_preenchida.xlsx` com a classificação das defesas e uma aba de auditoria. A coleta inicial depende de conexão à internet e da disponibilidade das páginas consultadas.

Exemplo para 2025, a partir de `projeto_final_MO810/`:

```bash
cd analises/02_defesas
python classificar_defesas.py --historico "Histórico de Defesas de Mestrado e Doutorado.xlsx" --template "Classificação.xlsx" --ano 2025 --saida "Classificação_2025_preenchida.xlsx"
```

Substitua o ano e o nome da saída para executar a coleta dos demais anos. Se as planilhas preenchidas e revisadas já estiverem disponíveis, **não as sobrescreva**: elas contêm correções manuais de nomes, afiliações e composição das bancas.

**Validação dos prêmios — `classificar_defesas_premios_validados.py`.** Atualiza as informações de prêmios das planilhas preenchidas utilizando exclusivamente o CSV validado da etapa 01. No modo `--atualizar-classificacao`, não realiza uma nova coleta na internet e preserva os dados das defesas corrigidos manualmente; apenas as informações de prêmios e os campos correspondentes da auditoria são atualizados.

Ainda dentro de `analises/02_defesas/`, execute:

```bash
python classificar_defesas_premios_validados.py --atualizar-classificacao "Classificação_2023_preenchida.xlsx" --premios-base "../01_premios/resultado/categorias_ic/classificacao_detalhada.csv" --ano 2023 --saida "Classificação_2023_premios_validados.xlsx"
python classificar_defesas_premios_validados.py --atualizar-classificacao "Classificação_2024_preenchida.xlsx" --premios-base "../01_premios/resultado/categorias_ic/classificacao_detalhada.csv" --ano 2024 --saida "Classificação_2024_premios_validados.xlsx"
python classificar_defesas_premios_validados.py --atualizar-classificacao "Classificação_2025_preenchida.xlsx" --premios-base "../01_premios/resultado/categorias_ic/classificacao_detalhada.csv" --ano 2025 --saida "Classificação_2025_premios_validados.xlsx"
cd ../..
```

Por padrão, a validação considera todo o período de premiações de 2019–2025, inclusive registros posteriores ao ano da defesa. Assim, os prêmios não representam necessariamente a situação do docente no momento da defesa. As planilhas anuais possuem 52 colunas; a etapa 03 utiliza a primeira aba de cada uma delas.

### 3.3. `analises/03_processamento` — Integração das bases

O script `analisar_docentes_ic.py` substitui a etapa anteriormente realizada em notebook. Integra:

- As três planilhas `Classificação_{ano}_premios_validados.xlsx`.
- O cadastro `Relação gênero e raça IC.xlsx`, localizado em `graficos/graficos_docents_IC_2026_site/`.
- A base `classificacao_detalhada.csv`, produzida na etapa 01.

O processamento unifica as defesas, identifica os docentes do IC, consolida suas participações por função e por defesa distinta e associa os registros de premiação por docente. As premiações são lidas diretamente da base validada, sem reclassificação por palavras-chave e sem depender da presença do docente em bancas. Os registros de gênero, raça/cor e vínculo são provenientes do cadastro do projeto.

A partir de `projeto_final_MO810/`:

```bash
python analises/03_processamento/analisar_docentes_ic.py
```

A raiz do projeto é localizada automaticamente a partir do próprio script. A execução não requer Jupyter Notebook e não modifica as planilhas de entrada. Os resultados são gravados em `analises/03_processamento/resultado_analise_validada/`:

| Arquivo | Conteúdo |
| --- | --- |
| `Classificação_2023_2025_unificada.xlsx` | Defesas dos três anos reunidas, com a coluna adicional `Ano`. |
| `tabela_docentes_ic.csv` | Consolidação por docente, incluindo vínculos pendentes de verificação. |
| `tabela_docentes_ic_filtrada.csv` | Base principal do site, excluindo os vínculos pendentes. |
| `tabela_docentes_ic_com_codigo.xlsx` | Planilha de resumo e auditoria do processamento. |
| `premios_docentes_detalhados.csv` | Registros da base validada de premiações. |
| `resumo_premios_por_categoria.csv` | Quantidades por categoria e situação de premiação. |

Os CSVs consolidados mantêm a estrutura de 17 colunas, incluindo as contagens de funções, bancas distintas, prêmios/menções, finalistas/indicações e outros reconhecimentos. Os arquivos CSV são gravados em UTF-8 com BOM, separados por ponto e vírgula.

### 3.4. `graficos` — Dados complementares e visualizações

Reúne os materiais de apoio às análises e às visualizações. Duas planilhas desta pasta são também fontes de entrada do site:

| Subdiretório | Fonte e finalidade |
| --- | --- |
| `graficos_total_defesas/` | `Anuário 2015-2025.xlsx`: informações históricas agregadas sobre o corpo docente, carreira, regimes de trabalho, sexo e raça/cor. |
| `graficos_docents_IC_2026_site/` | `Relação gênero e raça IC.xlsx`: cadastro nominal do levantamento do projeto, incluindo gênero, raça/cor e vínculo. |

As informações de 2026 vêm do levantamento da equipe e da [lista de docentes do IC](https://ic.unicamp.br/docentes/), não de uma edição de 2026 do Anuário Estatístico.

## 4. `site_ic` — Escolha sua banca

O site utiliza Python para preparar e fornecer os dados em JSON e HTML, CSS e JavaScript para as visualizações e interações. As fontes são lidas diretamente dos diretórios anteriores; não é necessário copiar ou renomear os arquivos de entrada.

### Fontes de entrada

Todos os caminhos abaixo são relativos a `projeto_final_MO810/`.

| Fonte | Caminho |
| --- | --- |
| Anuário | `graficos/graficos_total_defesas/Anuário 2015-2025.xlsx` |
| Premiações validadas | `analises/01_premios/resultado/categorias_ic/classificacao_detalhada.csv` |
| Docentes consolidados | `analises/03_processamento/resultado_analise_validada/tabela_docentes_ic_filtrada.csv` |
| Defesas unificadas | `analises/03_processamento/resultado_analise_validada/Classificação_2023_2025_unificada.xlsx` |
| Cadastro nominal | `graficos/graficos_docents_IC_2026_site/Relação gênero e raça IC.xlsx` |

### Scripts e arquivos da interface

| Arquivo | Responsabilidade |
| --- | --- |
| `caminhos.py` | Localiza a raiz do projeto e verifica as cinco fontes de entrada. |
| `preparar_dados.py` | Reconstrói as participações por docente, defesa e função; lê os reconhecimentos da base validada e confere os totais com a tabela consolidada. |
| `conferir_dados.py` | Executa testes de regressão e consistência, inclusive a comparação individual dos registros de premiação. |
| `site_escolha_sua_banca.py` | Integra as fontes, disponibiliza os dados em JSON e inicia o servidor HTTP local. |
| `testar_interface.py` | Testa os gráficos e seletores da interface em Chromium usando os dados produzidos pelo Python. |
| `index.html` | Estrutura, textos e seções da página. |
| `estilo.css` | Cores, tipografia, layout e adaptação a diferentes telas. |
| `interacao.js` | Gráficos, filtros, mapa de docentes, fichas, troféu e simulador. |
| `dados/participacoes.csv` | Registros derivados de participação por docente, ano, defesa e função. |
| `dados/reconhecimentos.csv` | Registros derivados de premiação com título, ano, categoria e situação. |

A pasta `site_ic/dados/` é criada automaticamente e contém apenas os arquivos derivados. Não coloque nela cópias das fontes originais.

### Preparação, validação e execução

A partir de `projeto_final_MO810/`:

```bash
python site_ic/caminhos.py
python site_ic/preparar_dados.py
python site_ic/conferir_dados.py
python site_ic/site_escolha_sua_banca.py
```

Acesse [http://localhost:8001](http://localhost:8001). Para escolher outra porta, use `--port 8002`, por exemplo. A inicialização do servidor também regenera os CSVs derivados; a execução separada dos scripts permite conferir cada etapa. Em caso de divergências, o processamento interrompe a execução para evitar apresentar dados inconsistentes. Após atualizar as fontes, reinicie o servidor.

### Teste opcional da interface

```bash
python -m pip install playwright
python -m playwright install chromium
python site_ic/testar_interface.py
```

O teste confere 45 combinações de gráficos no DOM do Chromium, com os arquivos reais de HTML, CSS e JavaScript e o JSON produzido pelo Python. O transporte da API é simulado, portanto o teste não exige um servidor HTTP em execução.

## 5. Recortes temporais e unidades de contagem

| Conjunto de dados | Período | Unidade |
| --- | --- | --- |
| Anuário Estatístico | 2015–2025 | Docentes ativos por ano. |
| Cadastro do projeto | 2026 | Docentes e situação no cadastro consultado. |
| Premiações e distinções | 2019–2025 | Registro distinto por docente, ano e título. |
| Defesas | 2023–2025 | Defesa, docente, par docente–defesa ou função. |

No recorte validado utilizado pelos testes atuais:

| Indicador | Total |
| --- | ---: |
| Docentes no cadastro filtrado | 85 |
| Docentes ativos no levantamento de 2026 | 47 |
| Defesas de mestrado | 106 |
| Defesas de doutorado | 67 |
| **Defesas analisadas** | **173** |
| Docentes distintos participantes | 59 |
| Pares distintos docente–defesa | 533 |
| Funções exercidas nas defesas | 705 |
| Docentes contemplados por reconhecimentos | 44 |
| Premiações e colocações | 213 |
| Menções honrosas | 36 |
| Finalistas e indicações | 16 |
| Outros reconhecimentos | 28 |
| **Registros de premiações e distinções** | **293** |
| Categorias de premiação | 28 |

Um docente que seja orientador e titular na mesma defesa ocupa duas funções, mas constitui apenas um par docente–defesa. Os gráficos agregados de bancas por gênero, raça/cor e vínculo acompanham o ano e a unidade de contagem selecionada; não acompanham os filtros nominais, de função ou de situação do mapa de docentes.

O subtotal de premiações/colocações e menções honrosas é de 249 registros, diferente do total de 293 reconhecimentos. As distribuições agregadas de premiações não acompanham os seletores individuais do troféu. Uma parte dos registros pode se referir a orientandos, trabalhos coletivos ou outras distinções associadas ao docente; não devem ser interpretados automaticamente como prêmios recebidos pessoalmente.

Os testes de regressão utilizam esses totais para detectar mudanças inesperadas. Se novas entradas ou anos forem acrescentados, revise os recortes, os nomes dos arquivos e os valores esperados nos testes antes de disponibilizar os resultados.

## 6. Fontes e cuidados metodológicos

- [Anuário Estatístico da Unicamp](https://www.aeplan.unicamp.br/anuarios/): dados agregados sobre docentes entre 2015 e 2025.
- [Anuário de Pesquisa — Prêmios e Homenagens](https://pesquisa.dados.unicamp.br/premios): exportações internas e externas de 2019 a 2025.
- [Instituto de Computação da Unicamp](https://ic.unicamp.br/): páginas de defesas e relação de docentes, complementadas pelas planilhas do projeto.

O desenvolvimento do site con