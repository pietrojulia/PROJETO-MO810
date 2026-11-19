"""Checks rápidos: execute python conferir_dados.py antes de publicar."""
from site_escolha_sua_banca import montar_dados
from collections import Counter

dados = montar_dados()
assert dados['metricas']['defesas']==173
assert dados['metricas']['pessoasCadastro']==95
assert dados['metricas']['participacoesTabela']==723
assert dados['metricas']['premiosTabela']==188
assert dados['docentes']['anos']['2025']['total']==44
assert dados['docentes']['anos']['2026']['total']==47
assert dados['docentes']['anos']['2026']['regimes']=={'Parcial':1,'Completo':0,'Integral':46}
assert len(dados['docentes']['anos'])==12
assert not any('genero' in p or 'raca' in p for p in dados['professores'])
assert sum(dados['distribuicoes']['bancas']['todos']['status'].values())==723
assert sum(dados['distribuicoes']['premios']['status'].values())==188
print('TESTES OK: métricas, 2015–2026, fontes e ausência de dados sensíveis por nome na API.')
print('Divergências para revisão:',dados['auditoria']['quantidadeDivergencias'])
for item in dados['auditoria']['divergencias']:
    print('-',item['professor'],item.get('diferencas',item.get('premios')))
