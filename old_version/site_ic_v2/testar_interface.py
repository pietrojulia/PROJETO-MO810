"""Teste opcional: pip install playwright; playwright install chromium.
Renderiza HTML, CSS e JS reais em Chromium, com fetch da API simulado pelo JSON
produzido pelo servidor Python a partir dos arquivos de dados reais.
"""
import json
import os
import re
from pathlib import Path
from playwright.sync_api import sync_playwright
from site_escolha_sua_banca import montar_dados

ROOT=Path(__file__).resolve().parent
data=montar_dados()
payload=json.dumps(data,ensure_ascii=False).replace('</','<\\/')
html=(ROOT/'index.html').read_text(encoding='utf-8')
html=re.sub(r'<link[^>]+(?:fonts\.googleapis|fonts\.gstatic)[^>]*>','',html)
html=re.sub(r'<link rel="stylesheet"[^>]*>',lambda _: '<style>'+(ROOT/'estilo.css').read_text(encoding='utf-8')+'</style>',html)
html=re.sub(r'<script defer src="[^"]+"></script>',lambda _: '<script>window.fetch=async()=>({ok:true,json:async()=>('+payload+')});</script><script>'+(ROOT/'interacao.js').read_text(encoding='utf-8')+'</script>',html)
with sync_playwright() as w:
    browser=w.chromium.launch(executable_path=os.environ.get('CHROMIUM_PATH'),args=['--no-sandbox'])
    page=browser.new_page(viewport={'width':1440,'height':1000})
    errors=[]
    page.on('pageerror',lambda e: errors.append(str(e)))
    page.on('console',lambda m: errors.append(m.text) if m.type=='error' else None)
    page.set_content(html,wait_until='load')
    page.wait_for_selector('#banca-dist-genero .bar-item')
    def conferir(prefix, expected):
        for dim in ('genero','raca','status'):
            actual={row.locator('.text').inner_text():int(row.locator('strong').inner_text().replace('.','')) for row in page.locator(f'#{prefix}-dist-{dim} .bar-item').all()}
            assert actual=={k:v for k,v in expected[dim].items() if v>0},(prefix,dim,actual,expected[dim])
    for ano in ('todos','2023','2024','2025'):
        page.select_option('#banca-ano',ano)
        for modo in ('docentes','participacoes','funcoes'):
            page.select_option('#banca-unidade',modo)
            conferir('banca',data['distribuicoes']['bancas'][ano][modo])
    for modo in ('registros','docentes','conquistas'):
        page.select_option('#premio-unidade',modo)
        conferir('premio',data['distribuicoes']['premios'][modo])
    page.select_option('#banca-ano','todos');page.select_option('#banca-unidade','docentes');page.select_option('#premio-unidade','registros')
    page.wait_for_timeout(700)
    assert not errors, errors
    browser.close()
print('INTERFACE OK: 45 gráficos conferidos no DOM, todos os anos e unidades de contagem; nenhum erro de JavaScript.')
