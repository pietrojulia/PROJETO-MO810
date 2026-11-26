#!/usr/bin/env python3
"""Exporta Escolha sua banca para GitHub Pages. Execute na pasta do site V2.
Não altera HTML, JS ou dados de origem. Publica apenas arquivos estáticos.
"""
from pathlib import Path
import argparse
import copy
import json
import runpy
import shutil

RAIZ = Path(__file__).resolve().parent
AVISO = 'Distribuição não publicada: classificações pessoais aguardam validação para divulgação pública.'


def trocar(texto, antigo, novo):
    if antigo not in texto:
        raise ValueError(f'Interface diferente da V2 esperada: {antigo[:80]}')
    return texto.replace(antigo, novo, 1)


def exportar(destino):
    from site_escolha_sua_banca import montar_dados
    # Valida a mesma base usada pela interface original; nenhuma nova classificação.
    runpy.run_path(str(RAIZ / 'conferir_dados.py'))
    dados = copy.deepcopy(montar_dados())
    if dados['auditoria']['quantidadeDivergencias']:
        raise ValueError('Há divergências na base; publicação interrompida.')
    for p in dados['professores']:
        for campo in ('genero', 'raca', 'origemDemografia', 'observacoes'):
            p.pop(campo, None)
        for participacao in p['participacoes']:
            participacao.pop('aluno', None)  # não usado pela interface
    # Não publica agregados de classificações nominais não verificadas.
    # Conserva os agregados oficiais do Anuário (2015–2025).
    d = dados['docentes']['anos']['2026']
    d['genero'] = {}
    d['raca'] = {}
    d['observacoes'] = [o for o in d['observacoes'] if 'Gênero e raça/cor' not in o] + [AVISO]
    for grupo in [*dados['distribuicoes']['bancas'].values(), dados['distribuicoes']['premios']]:
        for v in [grupo, *(x for x in grupo.values() if isinstance(x, dict) and 'genero' in x)]:
            v['genero'] = {}
            v['raca'] = {}
    dados['publicacaoEstatica'] = True
    # Somente arquivos explicitamente permitidos são exportados.
    html = (RAIZ / 'index.html').read_text(encoding='utf-8')
    js = (RAIZ / 'interacao.js').read_text(encoding='utf-8')
    html = trocar(html, 'estilo.css?v=20261125-v2', 'estilo.css?v=publicacao-1')
    html = trocar(html, 'interacao.js?v=20261125-v2', 'interacao.js?v=publicacao-1')
    html = trocar(html, '<br>VERSÃO 2', '')
    html = trocar(html, 'Por isso, mostramos apenas distribuições agregadas e não as exibimos nas fichas nominais. Categorias pendentes, como Pardo?, não equivalem a autodeclaração. Grupos muito pequenos podem permitir reidentificação; revise os agregados antes de publicação pública.', 'Na versão pública, preservamos os agregados oficiais do Anuário e reservamos as distribuições do levantamento nominal para revisão da equipe. Categorias pendentes não equivalem a autodeclaração; mesmo pequenos agregados podem permitir reidentificação.')
    html = trocar(html, 'o servidor interrompe a carga quando há divergências entre os detalhes, a tabela consolidada e a base validada de premiações.', 'a exportação dos dados é interrompida quando há divergências entre os detalhes, a tabela consolidada e a base validada de premiações.')
    html = trocar(html, 'A seleção individual e o ano do troféu não alteram estes gráficos, para proteger as classificações pessoais não confirmadas.', 'A seleção individual e o ano do troféu não alteram estes gráficos. As distribuições de gênero e raça/cor do levantamento aguardam validação para publicação.')
    html = trocar(html, 'Os gráficos acompanham apenas o ano, não os filtros nominais, de função ou status.', 'Os gráficos acompanham apenas o ano, não os filtros nominais, de função ou status. As distribuições de gênero e raça/cor do levantamento aguardam validação para publicação.')
    js = trocar(js, "fetch('/api/dados',{cache:'no-store'})", "fetch('./dados_publicos.json',{cache:'no-store'})")
    js = trocar(js, 'Versão incompatível do servidor. Encerre a versão antiga e execute o servidor do novo pacote na porta 8001', 'Versão incompatível dos dados. Exporte novamente a pasta de publicação')
    js = trocar(js, 'Confira o terminal do VS Code e as planilhas na pasta dados.', 'Confira se os arquivos estáticos foram publicados juntos.')
    # Mantém os cartões visíveis, distinguindo supressão de ausência de dados.
    js = trocar(js, "  bars('banca-dist-raca',agg.raca);", "  bars('banca-dist-raca',agg.raca);\n  for(const id of ['banca-dist-genero','banca-dist-raca']) $(id).innerHTML='<p class=\"small-note\">"+AVISO+"</p>';")
    js = trocar(js, "bars('premio-dist-status',agg.status,{colors:STATUS_COLORS});", "bars('premio-dist-status',agg.status,{colors:STATUS_COLORS});\n  for(const id of ['premio-dist-genero','premio-dist-raca']) $(id).innerHTML='<p class=\"small-note\">"+AVISO+"</p>';")
    js = trocar(js, "  renderTimeline();\n}", "  if(DOC_ANO==='2026'){\n    for(const id of ['doc-genero','doc-raca']) $(id).innerHTML='<p class=\"small-note\">"+AVISO+"</p>';\n    $('doc-genero-nota').textContent=''; $('doc-raca-nota').textContent='';\n  }\n  renderTimeline();\n}")
    js = trocar(js, 'M no Anuário / homem cis no levantamento de 2026', 'M no Anuário')
    js = trocar(js, 'F no Anuário / mulher cis no levantamento de 2026', 'F no Anuário')
    js = trocar(js, "      ['Pardo? (2026, a conferir)',RACE_COLORS['Pardo?']],\n", '')
    js = trocar(js, 'De 2015 a 2025: sexo tabulado M/F do Anuário. Em 2026: categorias de gênero do levantamento. As medidas não são diretamente equivalentes.', 'De 2015 a 2025: sexo tabulado M/F do Anuário. As classificações de 2026 aguardam validação para publicação.')
    js = trocar(js, 'Raça/cor não está disponível no Anuário de 2015 a 2022. O retrato de 2026 contém classificações pendentes e não constitui uma continuação oficial da série.', 'Raça/cor não está disponível no Anuário de 2015 a 2022. As classificações de 2026 aguardam validação para publicação.')
    destino = destino.resolve()
    if destino == RAIZ or RAIZ / 'dados' == destino or (RAIZ / 'dados') in destino.parents:
        raise ValueError('Escolha uma pasta de saída separada dos arquivos de origem.')
    destino.mkdir(parents=True, exist_ok=True)
    permitidos = {'index.html','estilo.css','interacao.js','dados_publicos.json','.nojekyll'}
    extras = {x.name for x in destino.iterdir()} - permitidos
    if extras:
        raise ValueError(f'Pasta de publicação contém arquivos adicionais: {sorted(extras)}. Use uma pasta vazia.')
    (destino / 'index.html').write_text(html,encoding='utf-8')
    (destino / 'interacao.js').write_text(js,encoding='utf-8')
    shutil.copyfile(RAIZ/'estilo.css',destino/'estilo.css')
    (destino/'dados_publicos.json').write_text(json.dumps(dados,ensure_ascii=False,allow_nan=False,separators=(',',':')),encoding='utf-8')
    (destino/'.nojekyll').touch()
    print(f'Exportação concluída: {destino}\nPublique somente os cinco arquivos dessa pasta, após revisão da equipe.')


if __name__ == '__main__':
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--saida',type=Path,default=RAIZ/'publicacao')
    exportar(p.parse_args().saida)
