# Versão 2 — correção dos gráficos vazios

Encerre o servidor antigo (Ctrl+C), extraia o ZIP em uma pasta nova e abra essa pasta no VS Code. Não misture arquivos de versões diferentes. Execute `python conferir_dados.py` e `python site_escolha_sua_banca.py --port 8001`. Abra http://localhost:8001 e pressione Ctrl+F5. A interface mostra VERSÃO 2 no cabeçalho. O servidor envia cabeçalhos que desabilitam cache e o HTML referencia uma versão nova dos arquivos JS/CSS.

A resposta JSON oferece as medidas nomeadas para a interface nova e, para evitar gráficos vazios na interface antiga, também oferece as distribuições padrão diretamente (docentes distintos nas bancas e todos os registros nos prêmios). Os textos e seletores corretos requerem o HTML e JS novos. A interface nova verifica a versão do servidor antes de renderizar.

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
| `dados/tabela_docentes_ic_filtrada.csv` | **Tabela consolidada principal**, com totais e status de cada docente |
| `dados/Classificação_2023_2025_unificada.xlsx` | Fonte de defesas para detalhar participações por ano |
| `dados/Anuario_2015_2025.xlsx` | Fonte dos dados agregados de 2015 a 2025 |
| `dados/Relacao_genero_raca_IC.xlsx` | Cadastro nominal e categorias registradas pelo projeto, usado para 2026 |
| `dados/participacoes.csv` | Detalhes de participação gerados pelo script |
| `dados/reconhecimentos.csv` | Registros oficiais derivados da base validada, com categoria e situação |
| `dados/classificacao_detalhada.csv` | Fonte única dos 293 registros de premiação (2019–2025) |

As entradas originais ficam na pasta `dados`. O servidor busca somente nesses nomes fixos. Não é necessário alterar caminhos no código se a pasta for mantida.

## 4. Fontes e limites, para deixar claros no site

- **2015–2025:** [Anuário Estatístico da Unicamp](https://www.aeplan.unicamp.br/anuarios/), abas `Docentes Ativos`, `Magistério`, `Sexo` e `Raça`.
- **2026:** levantamento do projeto com base na [lista de docentes do IC](https://ic.unicamp.br/docentes/) e na relação nominal da equipe, consultadas em **19 de novembro de 2026**. O total de 47 docentes ativos não deve ser apresentado como dado do Anuário 2026.
- **2025:** o Anuário registra 44 docentes, dos quais 43 em regime integral e 1 em completo (MS3.1, M). A aba `Sexo` soma **43** registros enquanto a aba de ativos soma **44**. Mantemos a divergência explicitada, sem inventar a categoria do registro ausente.
- **2026:** a relação nominal tem 47 ativos, MS3: 17, MS5: 18, MS6: 12; o levantamento indica 46 em regime integral e 1 parcial (MS5.1, M). A mesma relação classifica 40 como homem cis, 7 como mulher cis, 42 como branco, 1 como preto, 1 como pardo e **3 como `Pardo?`**, pendentes de confirmação.
- **Status atual, não histórico:** aposentado, ativo, ex-professor, colaborador e in memoriam são as etiquetas do cadastro usado no projeto. Elas **não** são o status da pessoa no momento da defesa.
- **Gênero e raça individuais:** a coluna `Lattes` da relação registra **84 de 85 casos como sem informação pública de gênero e raça**. Esses rótulos são classificações do levantamento, não autodeclarações verificadas. Por isso, o JSON da página não envia gênero/raça junto dos nomes e os gráficos agregados não mudam com busca por nome, papel ou status. Não publique identidades sensíveis presumidas sem validação documental e reflexão ética com a equipe.
- **Prêmios:** o inventário validado de 2019–2025 é lido diretamente de `classificacao_detalhada.csv`; não se extraem prêmios das bancas nem se reclassificam títulos. Não trate o inventário como uma contagem definitiva de prêmios recebidos pessoalmente pelo professor; alguns registros se referem a orientandos, trabalho coletivo ou outra distinção.


## 3. Atualização e validação dos dados

- Copie `pre_proc/resultado_analise_validada/tabela_docentes_ic_filtrada.csv` para `dados/tabela_docentes_ic_filtrada.csv`. O arquivo enviado nesta conversa era a versão anterior (703 funções); o pacote usa a versão validada (705).
- Copie `premios/resultado/categorias_ic/classificacao_detalhada.csv` para `dados/classificacao_detalhada.csv`. Esta é uma cópia da fonte oficial; atualize-a sempre junto com o CSV do notebook.
- As planilhas Excel enviadas permanecem intactas.
- Execute `python preparar_dados.py` e `python conferir_dados.py`. Os testes de regressão esperam os totais do recorte atual. Em novos anos, revise conscientemente esses totais.
- Execute `python site_escolha_sua_banca.py`. A inicialização regenera os arquivos derivados e interrompe em caso de divergência. Reinicie o servidor após atualizar os dados.

## Unidades de contagem

Bancas (2023–2025): 173 defesas, 59 docentes distintos participantes, 533 pares distintos docente–defesa, 705 funções. Um orientador que também seja titular conta duas funções, mas somente um par docente–defesa. Os gráficos agregados por perfil acompanham o ano, não filtros nominais, status ou papel.

Prêmios (2019–2025): 293 registros distintos para 44 docentes, incluindo 213 premiações/colocações, 36 menções honrosas, 16 finalistas/indicações e 28 outros reconhecimentos. O subtotal de conquistas é 249. As distribuições são globais e não acompanham os seletores individuais do troféu.

**Uso local:** a API não envia gênero ou raça individual. Entretanto, pequenas contagens agregadas e dados nominais de participação/prêmios podem permitir inferências. Revise a divulgação e a proteção de pequenos grupos antes de publicar; o pacote não é uma versão pública anonimizada. O servidor disponibiliza somente os três arquivos de interface e a API, não a pasta dados.

## Teste opcional da interface

Instale `pip install playwright` e `playwright install chromium`, depois execute `python testar_interface.py`. O teste usa o JSON real produzido pelo Python e o HTML/JS/CSS reais, com transporte de API simulado, e confere o conteúdo de 45 gráficos em Chromium. Não exige um servidor HTTP em execução.
