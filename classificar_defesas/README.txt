CLASSIFICAÇÃO DAS DEFESAS DO IC/UNICAMP — PACOTE ATUALIZADO
=============================================================

Conteúdo do pacote
------------------
- classificar_defesas.py
  Script com consulta a todos os anos disponíveis de prêmios e banca.

- requirements_classificar_defesas.txt
  Dependências Python.

- Histórico de Defesas de Mestrado e Doutorado.xlsx
  Arquivo de entrada enviado.

- Classificação.xlsx
  Modelo de classificação.


INSTALAÇÃO NO UBUNTU/WSL
------------------------
Entre na pasta descompactada e execute:

sudo apt update
sudo apt install -y python3-pip python3-venv
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements_classificar_defesas.txt

EXECUTAR 2023
-------------
python3 classificar_defesas.py --historico "Histórico de Defesas de Mestrado e Doutorado.xlsx" --template "Classificação.xlsx" --ano 2023 --saida "Classificação_2023_preenchida.xlsx"

EXECUTAR 2024
-------------
python3 classificar_defesas.py --historico "Histórico de Defesas de Mestrado e Doutorado.xlsx" --template "Classificação.xlsx" --ano 2024 --saida "Classificação_2024_preenchida.xlsx"

EXECUTAR 2025
-------------
python3 classificar_defesas.py --historico "Histórico de Defesas de Mestrado e Doutorado.xlsx" --template "Classificação.xlsx" --ano 2025 --saida "Classificação_2025_preenchida.xlsx"

EXECUTAR 2023, 2024 E 2025 EM SEQUÊNCIA
---------------------------------------
for ano in 2023 2024 2025; do
  python3 classificar_defesas.py \
    --historico "Histórico de Defesas de Mestrado e Doutorado.xlsx" \
    --template "Classificação.xlsx" \
    --ano "$ano" \
    --saida "Classificação_${ano}_preenchida.xlsx"
done


PRÊMIOS
-------
Por padrão, o script descobre e consulta todas as edições disponíveis no
Anuário de Pesquisa da Unicamp. O ano passado em --ano é o ano DAS DEFESAS.

SUPLENTES
---------
A saída possui Suplente 1, Suplente 2 e Suplente 3 para todas as defesas.
Quando não houver terceiro suplente, as respectivas células ficam em branco.

ATUALIZAÇÕES DE PRÊMIOS
------------------------
Após validação dos prêmio e reconhecimentos externos e internos, sugiro realizar a
atualização das planilhas de classificação. Isso pois, o primeiro script foi feito 
concomitantemente a classificação das defesas e a busca dos prêmios. 

python3 classificar_defesas_premios_coerentes.py \
  --atualizar-classificacao "Classificação_2023_preenchida(4).xlsx" \
  --premios-base premios_ic_validados.csv \
  --ano 2023 \
  --saida "Classificação_2023_premios_coerentes.xlsx"