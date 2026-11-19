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

3. Abra `http://localhost:8000` se o navegador não abrir automaticamente. Pare o servidor com **Ctrl+C**.


## 2. Organização do projeto

| Arquivo | Responsabilidade |
| --- | --- |
| `index.html` | Textos, ordem das seções, títulos e estrutura |
| `estilo.css` | Cores, tipografia, espaços e versão para celular |
| `interacao.js` | Seletores, mapa de pessoas, gráficos, troféu e simulador |
| `site_escolha_sua_banca.py` | Servidor local e criação dos dados JSON |
| `preparar_dados.py` | Reconstrói os arquivos detalhados de participações e prêmios a partir do Excel |
| `dados/tabela_docentes_ic_filtrada.csv` | **Tabela consolidada principal**, com totais e status de cada docente |
| `dados/Classificação_2023_2025_unificada.xlsx` | Fonte de defesas para detalhar participações por ano |
| `dados/Anuario_2015_2025.xlsx` | Fonte dos dados agregados de 2015 a 2025 |
| `dados/Relacao_genero_raca_IC.xlsx` | Cadastro nominal e categorias registradas pelo projeto, usado para 2026 |
| `dados/participacoes.csv` | Detalhes de participação gerados pelo script |
| `dados/reconhecimentos.csv` | Registros distintos de reconhecimento gerados pelo script |

As entradas originais ficam na pasta `dados`. O servidor busca somente nesses nomes fixos. Não é necessário alterar caminhos no código se a pasta for mantida.

## 4. Fontes e limites, para deixar claros no site

- **2015–2025:** [Anuário Estatístico da Unicamp](https://www.aeplan.unicamp.br/anuarios/), abas `Docentes Ativos`, `Magistério`, `Sexo` e `Raça`.
- **2026:** levantamento do projeto com base na [lista de docentes do IC](https://ic.unicamp.br/docentes/) e na relação nominal da equipe, consultadas em **19 de novembro de 2026**. O total de 47 docentes ativos não deve ser apresentado como dado do Anuário 2026.
- **2025:** o Anuário registra 44 docentes, dos quais 43 em regime integral e 1 em completo (MS3.1, M). A aba `Sexo` soma **43** registros enquanto a aba de ativos soma **44**. Mantemos a divergência explicitada, sem inventar a categoria do registro ausente.
- **2026:** a relação nominal tem 47 ativos, MS3: 17, MS5: 18, MS6: 12; o levantamento indica 46 em regime integral e 1 parcial (MS5.1, M). A mesma relação classifica 40 como homem cis, 7 como mulher cis, 42 como branco, 1 como preto, 1 como pardo e **3 como `Pardo?`**, pendentes de confirmação.
- **Status atual, não histórico:** aposentado, ativo, ex-professor, colaborador e in memoriam são as etiquetas do cadastro usado no projeto. Elas **não** são o status da pessoa no momento da defesa.
- **Gênero e raça individuais:** a coluna `Lattes` da relação registra **84 de 85 casos como sem informação pública de gênero e raça**. Esses rótulos são classificações do levantamento, não autodeclarações verificadas. Por isso, o JSON da página não envia gênero/raça junto dos nomes e os gráficos agregados não mudam com busca por nome, papel ou status. Não publique identidades sensíveis presumidas sem validação documental e reflexão ética com a equipe.
- **Prêmios:** as descrições repetidas foram deduplicadas por pessoa, ano e título e classificadas automaticamente. Não trate o inventário como uma contagem definitiva de prêmios recebidos pessoalmente pelo professor; alguns registros se referem a orientandos, trabalho coletivo ou outra distinção.

