"""Testes de regressão para o recorte validado 2019–2025 / defesas 2023–2025."""
from collections import Counter
from site_escolha_sua_banca import montar_dados
from preparar_dados import premios_oficiais, chave, normalizar

x = montar_dados()
m = x['metricas']
for k,v in {'defesas':173,'pessoasCadastro':85,'participacoesTabela':705,
            'participacoesDistintas':533,'docentesParticipantes':59,
            'registrosPremios':293,'docentesPremiados':44,'premiosTabela':249}.items():
    assert m[k]==v, (k,m[k],v)
assert m['situacoesPremios']=={'Premiação / colocação':213,'Menção honrosa':36,'Finalista / indicação':16,'Outro reconhecimento':28}
assert len(m['categoriasPremios'])==28
assert sum(m['categoriasPremios'].values())==293
assert len(x['docentes']['anos'])==12
assert x['docentes']['anos']['2025']['total']==44
assert x['docentes']['anos']['2026']['total']==47
assert x['docentes']['anos']['2026']['regimes']=={'Parcial':1,'Completo':0,'Integral':46}
assert not any('genero' in p or 'raca' in p or 'origemDemografia' in p for p in x['professores'])
assert not x['auditoria']['divergencias']
assert x['auditoria']['quantidadeDivergencias']==0
# Igualdade registro a registro, incluindo categoria e situação, não só totais.
orig = Counter((k,v['titulo'],v['categoria'],v['situacao']) for k,v in premios_oficiais().items())
api = Counter(((p['id'],r['ano'],normalizar(r['titulo'])),r['titulo'],r['categoria'],r['situacao']) for p in x['professores'] for r in p['premios'])
assert orig==api
for unidade,total in [('docentes',59),('participacoes',533),('funcoes',705)]:
    for dist in x['distribuicoes']['bancas']['todos'][unidade].values(): assert sum(dist.values())==total
for unidade,total in [('docentes',44),('registros',293),('conquistas',249)]:
    for dist in x['distribuicoes']['premios'][unidade].values(): assert sum(dist.values())==total
for ano in ('2023','2024','2025'):
    for unidade in ('docentes','participacoes','funcoes'):
        valores=[sum(d.values()) for d in x['distribuicoes']['bancas'][ano][unidade].values()]
        assert len(set(valores))==1
assert all(sum(x['distribuicoes']['bancas'][a]['funcoes']['status'].values())>0 for a in ('2023','2024','2025'))
print('TESTES OK: 293 registros oficiais idênticos, 28 categorias, 705 funções, 533 pares docente–defesa, 59 participantes, 44 contemplados; sem divergências.')

assert x['versaoInterface']=='2026-11-25-v2'
for a in ('todos','2023','2024','2025'):
    for dim in ('genero','raca','status'):
        assert x['distribuicoes']['bancas'][a][dim]==x['distribuicoes']['bancas'][a]['docentes'][dim]
for dim in ('genero','raca','status'):
    assert x['distribuicoes']['premios'][dim]==x['distribuicoes']['premios']['registros'][dim]
print('Compatibilidade do JSON e versão: OK.')
