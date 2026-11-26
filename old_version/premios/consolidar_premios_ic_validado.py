#!/usr/bin/env python3
"""Consolida as 14 exportações do Anuário Unicamp (2019–2025).
Executar: python consolidar_premios_ic.py --entrada . --saida resultado
"""
import argparse
import collections
from collections import Counter
import csv
import difflib
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

URL_ANUARIO = "https://pesquisa.dados.unicamp.br/premios"
URL_DOCENTES = "https://ic.unicamp.br/docentes/"
DOCENTES = """Alexandre Xavier Falcão
Allan Mariano de Souza
Anderson de Rezende Rocha
André Santanchè
Breno Bernard Nicolau de França
Bruno Barbieri de Pontes Cafeo
Carlos Alberto Astudillo Trujillo
Claudia Maria Bauzer Medeiros
Edmundo Roberto Mauro Madeira
Edson Borin
Emanuel Felipe Duarte
Esther Luna Colombini
Fábio Luiz Usberti
Flávio Keidi Miyazawa
Gabriel Capiteli Bertocco
Helena de Almeida Maia
Hélio Pedrini
Hervé Cédric Yviquel
Hilder Vitor Lima Pereira
Isaías Bittencourt Felzmann
Islene Calciolari Garcia
Juliana Freitag Borin
Julio Cesar dos Reis
Julio César López Hernández
Leandro Aparecido Villas
Lehilton Lelis Chaves Pedrosa
Lucas Francisco Wanner
Luiz Fernando Bittencourt
Nelson Luis Saldanha da Fonseca
Pedro Jussieu de Rezende
Rafael Crivellari Saliba Schouery
Ricardo Dahab
Rodolfo Jardim de Azevedo
Rubén Interian Kovaliova
Sandra Eliza Fontes de Avila
Santiago Valdés Ravelo
Zanoni Dias
Cid Carvalho de Souza
Guido Costa Souza de Araújo
Jacques Wainer
Paulo Lício de Geus
Maria Cecília Calani Baranauskas
Ricardo de Oliveira Anido
Diego de Freitas Aranha
Ricardo da Silva Torres
Cecília Mary Fischer Rubira""".splitlines()

INGRESSO = {
    "Allan Mariano de Souza": (2023, "https://www.sg.unicamp.br/visualizar-impressao/cad/deliberacoes/2023/392/0eb092abab36393ac1ec3db0105b6fcd"),
    "Carlos Alberto Astudillo Trujillo": (2023, "https://orcid.org/0000-0002-2407-0981"),
    "Emanuel Felipe Duarte": (2023, "https://www.sg.unicamp.br/cad/deliberacoes?ano=&descricao=&numero=&pagina=210&tipo="),
    "Bruno Barbieri de Pontes Cafeo": (2023, "https://www.escavador.com/diarios/5647359/DOESP/executivo-ii/2023-04-26?page=66"),
    "Rubén Interian Kovaliova": (2023, "https://www.sg.unicamp.br/cad/deliberacoes?ano=&descricao=&numero=&pagina=189&tipo="),
    "Hilder Vitor Lima Pereira": (2023, "https://www.sg.unicamp.br/cad/deliberacoes?ano=&descricao=298&numero=&pagina=2&tipo="),
    "Isaías Bittencourt Felzmann": (2024, "https://www.sg.unicamp.br/cad/deliberacoes?ano=&descricao=12&numero=&pagina=48&tipo="),
    "Gabriel Capiteli Bertocco": (2026, "https://www.sg.unicamp.br/cad/deliberacoes?ano=&descricao=computa%C3%A7%C3%A3o&numero=&pagina=3&tipo="),
    "Helena de Almeida Maia": (2026, "https://portal.dados.unicamp.br/perfil?docente=333034&nome_programa=&nome_unidade=INSTITUTO+DE+COMPUTACAO&origem=&sigla_unidade=IC"),
}
EX_DOCENTES = {"Diego de Freitas Aranha", "Ricardo da Silva Torres"}

def texto(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()

def normal(s):
    s = unicodedata.normalize("NFKD", texto(s)).casefold()
    return "".join(c for c in s if not unicodedata.combining(c))

def nome_key(s):
    return re.sub(r"^(?:prof(?:a)?\.?\s*(?:dr(?:a)?\.?\s*)?)", "", normal(s)).strip()

ROSTER = {nome_key(n): n for n in DOCENTES}
ROSTER[nome_key("Gabriel Bertocco")] = "Gabriel Capiteli Bertocco"

def premio_key(s, ano, nome):
    v = normal(s)
    if "zeferino vaz" in v:
        return "institucional: zeferino vaz"
    if "dedicacao ao ensino de graduacao" in v:
        return "institucional: dedicacao ao ensino"
    if ("proec" in v and ("premio" in v or "extensao" in v)) or v == "premio de extensao universitaria":
        return "institucional: proec"
    # Variantes claras do mesmo reconhecimento, conferidas nos arquivos.
    if ano == 2022 and nome == "Luiz Fernando Bittencourt" and v in {"best reviewer award noms 2022", "reviewer award noms 2022"}:
        return "reviewer award noms 2022"
    if ano == 2023 and nome == "Hervé Cédric Yviquel" and "erad 23, simplifying hpc application development with openmp cluster and unified memory" in v:
        return "erad 23: mencao honrosa simplifying hpc application development with openmp cluster and unified memory"

    return v

NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

def ler_xlsx(path):
    """Lê a primeira aba de exportação XLSX via XML, inclusive shared strings."""
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            shared = ["".join(t.text or "" for t in si.findall(".//x:t", NS))
                      for si in root.findall("x:si", NS)]
        root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
        for row in root.findall(".//x:sheetData/x:row", NS):
            vals = {}
            for c in row.findall("x:c", NS):
                col = re.match(r"[A-Z]+", c.attrib["r"])[0]
                v = c.findtext("x:v", default="", namespaces=NS)
                if c.get("t") == "s":
                    v = shared[int(v)]
                elif c.get("t") == "inlineStr":
                    v = "".join(t.text or "" for t in c.findall(".//x:t", NS))
                vals[col] = v
            yield int(row.attrib["r"]), vals

def coletar(entrada):
    raw = []
    combos = set()
    for path in sorted(entrada.glob("Tabela de Prêmios e Homenagens*.xlsx")):
        m = re.search(r"(20\d{2})", path.stem)
        if not m or not 2019 <= int(m[1]) <= 2025:
            continue
        ano = int(m[1])
        tipo = "Interno" if "Internos" in path.name else "Externo" if "Externos" in path.name else None
        if not tipo:
            continue
        if (ano,tipo) in combos:
            raise ValueError(f"Exportação duplicada para {ano} {tipo}: {path.name}")
        combos.add((ano,tipo))
        rows = iter(ler_xlsx(path))
        _, headers = next(rows)
        cols = {texto(v): k for k, v in headers.items()}
        unidade = cols.get("Sigla") or cols.get("Unidade")
        for linha, r in rows:
            if texto(r.get(unidade)).upper() != "IC":
                continue
            if "Ano" in cols and int(r[cols["Ano"]]) != ano:
                raise ValueError(f"Ano divergente em {path.name}, linha {linha}")
            nome = texto(r.get(cols["Nome"]))
            premio = texto(r.get(cols["Prêmio"]))
            raw.append({"ano": ano, "tipo": tipo, "nome": nome, "premio": premio,
                        "arquivo": path.name, "linha": linha})
    esperados = {(a,t) for a in range(2019,2026) for t in ("Interno","Externo")}
    if combos != esperados:
        raise ValueError(f"Faltam exportações: {sorted(esperados-combos)}")
    return raw

def consolidar(raw):
    grupos = collections.defaultdict(list)
    for r in raw:
        n = ROSTER.get(nome_key(r["nome"]), r["nome"])
        k = (r["ano"], nome_key(n), premio_key(r["premio"], r["ano"], n))
        grupos[k].append(r)
    dados = []
    for (ano, nk, pk), rs in grupos.items():
        base = sorted(rs, key=lambda r: (r["tipo"] != "Interno", r["arquivo"], r["linha"]))[0]
        nome = ROSTER.get(nome_key(base["nome"]), base["nome"])
        fonte = URL_DOCENTES if nome_key(nome) in ROSTER else ""
        status = "Docente identificado"
        nota = "Nome aceito na lista de docentes; vínculo histórico não verificado individualmente."
        if nome_key(nome) not in ROSTER:
            status = "Nome não localizado na lista docente"
            nota = "Excluído conforme validação do revisor: não é docente do IC no recorte."
        elif nome in INGRESSO:
            ingresso, fonte = INGRESSO[nome]
            if ano < ingresso:
                status = "Anterior à nomeação conhecida"
                nota = f"Nomeação/ingresso localizado: {ingresso}; exclusão confirmada pelo revisor."
            elif ano == ingresso:
                status = "Docente identificado"
                nota = f"Ano de ingresso: {ingresso}; inclusão no mesmo ano aprovada pelo revisor."
            else:
                nota = f"Posterior ao ano de nomeação/ingresso localizado ({ingresso})."
        elif nome in EX_DOCENTES:
            nota = "Ex-docente listado no site oficial; confirmar vínculo na data do prêmio."
        tipo = base["tipo"]
        origens = " + ".join(sorted({r["tipo"] for r in rs}))
        dados.append({"Ano":ano, "Professor / nome":nome, "Prêmio":base["premio"],
                      "Tipo":tipo, "Tabelas de origem":origens, "Situação":status,
                      "Observação de vínculo":nota, "Registros de origem":len(rs),
                      "Arquivo e linha":" | ".join(f'{r["arquivo"]} [linha {r["linha"]}]' for r in rs),
                      "Outras descrições":" | ".join(dict.fromkeys(r["premio"] for r in rs if r["premio"] != base["premio"])),
                      "Fonte docente":fonte, "Fonte prêmios":URL_ANUARIO})
    dados.sort(key=lambda r: (r["Ano"], nome_key(r["Professor / nome"]), normal(r["Prêmio"])))
    inter = collections.defaultdict(set)
    for r in dados:
        inter[(nome_key(r["Professor / nome"]),normal(r["Prêmio"]))].add(r["Ano"])
    for r in dados:
        anos = inter[(nome_key(r["Professor / nome"]),normal(r["Prêmio"]))]
        r["Mesmo texto em outros anos"] = ", ".join(str(a) for a in sorted(anos-{r["Ano"]}))
    por_pessoa_ano = collections.defaultdict(list)
    for r in dados:
        r["Descrição semelhante no ano"] = ""
        por_pessoa_ano[(r["Ano"], nome_key(r["Professor / nome"]))].append(r)
    for rs in por_pessoa_ano.values():
        for i, a in enumerate(rs):
            for b in rs[i+1:]:
                if difflib.SequenceMatcher(None,normal(a["Prêmio"]),normal(b["Prêmio"])).ratio() >= 0.86:
                    for x, y in ((a,b),(b,a)):
                        x["Descrição semelhante no ano"] = " | ".join(filter(None, [x["Descrição semelhante no ano"], y["Prêmio"]]))
    return dados

def separar(dados):
    principal = [r for r in dados if r["Situação"] in {"Docente identificado"}]
    fora = [r for r in dados if r["Situação"] not in {"Docente identificado"}]
    return principal, fora

def gravar_csv(path, rows):
    if not rows:
        return
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]), delimiter=";")
        w.writeheader()
        w.writerows(rows)


# CLASSIFICAÇÃO DAS PREMIAÇÕES

def normalizar(texto):
    texto = unicodedata.normalize('NFKC', texto).casefold()
    return ' '.join(texto.split())


def chave(linha):
    return tuple(normalizar(linha[c]) for c in ('Ano', 'Professor / nome', 'Prêmio'))



REGRAS = {
    ('2019', 'alexandre xavier falcão', 'melhor artigo (voto popular): altis: a fast and automatic lung and trachea ct-image segmentation method. medical physics, doi 10.1002/mp.13773, 46(11), pp. 4970--4982, nov 2019'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2019', 'alexandre xavier falcão', 'prêmio de inventores unicamp 2018 na categoria "patente concedida"'): ('Prêmio Inventores Unicamp', 'Premiação / colocação', ''),
    ('2019', 'alexandre xavier falcão', 'senior spie member'): ('Fellow / membro / senior member', 'Outro reconhecimento', ''),
    ('2019', 'anderson de rezende rocha', 'eleição - chair do ieee information forensics and security technical committee (ifs-tc)'): ('Eleição / nomeação para cargo ou comitê', 'Outro reconhecimento', ''),
    ('2019', 'anderson de rezende rocha', 'google research awards for latin america ? estudante: antonio theophilo (doutorado)'): ('Google Latin America Research Awards (LARA)', 'Premiação / colocação', ''),
    ('2019', 'anderson de rezende rocha', 'melhor tese de doutorado ? estudante: allan da silva pinto'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2019', 'anderson de rezende rocha', 'melhor tese de doutorado ? estudante: josé ramon trindade pires'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2019', 'anderson de rezende rocha', 'melhor tese de doutorado ? prêmio capes de teses ? estudante: allan da silva pinto'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2019', 'breno bernard nicolau de frança', 'best paper nominee - xxii iberoamerican conference on software engineering'): ('Best Paper Award / Melhor artigo', 'Finalista / indicação', ''),
    ('2019', 'breno bernard nicolau de frança', 'keynote speaker - i workshop de modelagem e simulação de sistemas intensivos em software (mssis)'): ('Keynote / distinguished lecturer', 'Outro reconhecimento', ''),
    ('2019', 'breno bernard nicolau de frança', 'melhor artigo do xx workshop de testes e tolerência a falhas do sbrc'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2019', 'cid carvalho de souza', '4th place - workshop on open problems and hard instance challenges'): ('Outros prêmios e distinções específicos', 'Premiação / colocação', ''),
    ('2019', 'cid carvalho de souza', 'segundo lugar no concurso de dissertações de mestrado obtido pelo mestrando natanael ramos (orientador),'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2019', 'claudia maria bauzer medeiros', 'membro titular'): ('Fellow / membro / senior member', 'Outro reconhecimento', ''),
    ('2019', 'diego de freitas aranha', 'orientador da melhor dissertaçäo de mestrado em arquitetura de computadores e computação de alto desempenho'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2019', 'diego de freitas aranha', 'prëmio inventores - patente concedida e tecnologia licenciada'): ('Prêmio Inventores Unicamp', 'Premiação / colocação', ''),
    ('2019', 'edmundo roberto mauro madeira', 'prêmio de reconhecimento docente pela dedicação ao ensino de graduação'): ('Reconhecimento pela Dedicação ao Ensino de Graduação', 'Premiação / colocação', ''),
    ('2019', 'edson borin', 'best paper: avaliação das estruturas de arquivo para processamento de dados sísmicos com alto desempenho na nuvem computacional'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2019', 'edson borin', 'co-orientador da dissertação de mestrado de antonio carlos guimarães junior. 1o lugar concurso de teses e dissertações do xx simpósio em sistemas computacionais de alto desempenho'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2019', 'edson borin', 'prêmio cepetro em pesquisa, desenvolvimento e inovação 2018 - primeiro lugar na categoria graduação. aluno: lucas de magalhães araujo. orientador: edson borin'): ('Iniciação científica / graduação', 'Premiação / colocação', ''),
    ('2019', 'edson borin', 'prêmio proec de extensão universitária'): ('Prêmio ProEC de Extensão Universitária', 'Premiação / colocação', ''),
    ('2019', 'fábio luiz usberti', '2o lugar entre os melhores trabalhos do li sbpo'): ('Melhor trabalho – nível/tipo não informado', 'Premiação / colocação', ''),
    ('2019', 'fábio luiz usberti', 'medalha de ouro na final brasileira de maratona de programação (atuando como técnico)'): ('Maratona de programação', 'Premiação / colocação', ''),
    ('2019', 'hélio pedrini', '1o lugar - dissertação de mestrado (aluno: marcos roberto e souza)'): ('Dissertações de mestrado', 'Premiação / colocação', 'Mesmo aluno e ano que outro registro; evento não especificado, mantido.'),
    ('2019', 'hélio pedrini', '3o lugar - melhores pôsteres. aluno: manuel alberto córdova neira. pelee-text: a tiny convolutional neural network for multi-oriented scene text detection.'): ('Best Poster Award / Melhor pôster', 'Premiação / colocação', ''),
    ('2019', 'hélio pedrini', 'melhor dissertação de mestrado (aluno: marcos roberto e souza)'): ('Dissertações de mestrado', 'Premiação / colocação', 'Mesmo aluno e ano que outro registro; evento não especificado, mantido.'),
    ('2019', 'hélio pedrini', 'menção honrosa - dissertação de mestrado (aluno: marcos roberto e souza; título: digital video stabilization: algorithms and evaluation)'): ('Dissertações de mestrado', 'Menção honrosa', ''),
    ('2019', 'hélio pedrini', 'menção honrosa - prêmio mercosul de ciência e tecnologia (edição 2017). categoria pesquisador sênior.'): ('Outros prêmios e distinções específicos', 'Menção honrosa', ''),
    ('2019', 'hélio pedrini', 'prêmio - excelência didática do instituto de computação'): ('Excelência Didática (IC)', 'Premiação / colocação', ''),
    ('2019', 'hélio pedrini', 'prêmio - franklin v. taylor memorial award. trabalho: audio-visual emotion recognition using a hybrid deep convolutional neural network based on census transform.'): ('Outros prêmios e distinções específicos', 'Premiação / colocação', ''),
    ('2019', 'hélio pedrini', 'prêmio capes de tese 2019 (aluno: allan da silva pinto)'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2019', 'jacques wainer', 'concurso de teses e dissertações sbc - 1o lugar - aluno: ramon pires (co-orientador)'): ('Teses/dissertações – nível não informado', 'Premiação / colocação', ''),
    ('2019', 'julio césar lópez hernández', 'melhor artigo curto, xix simpósio brasileiro em segurança da informação e de sistemas computacionais (sbseg2019).'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2019', 'julio césar lópez hernández', 'prêmio de inventores'): ('Prêmio Inventores Unicamp', 'Premiação / colocação', ''),
    ('2019', 'leandro aparecido villas', 'excelência acadêmica, instituto de computação, universidade estadual de campinas'): ('Excelência Acadêmica (IC)', 'Premiação / colocação', ''),
    ('2019', 'leandro aparecido villas', 'menção honrosa dissertação de mestrado no concurso de teses e dissertações do simpósio brasileiro de redes de computadores e sistemas distribuídos'): ('Dissertações de mestrado', 'Menção honrosa', ''),
    ('2019', 'leandro aparecido villas', 'menção honrosa do xxxvi simpósio brasileiro de redes de computadores e sistemas distribuídos, sociedade brasileira de computação'): ('Menção honrosa – objeto não especificado', 'Menção honrosa', ''),
    ('2019', 'leandro aparecido villas', 'prêmio de melhor artigo em inglês no simpósio brasileiro de pesquisa operacional'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2019', 'lehilton lelis chaves pedrosa', 'best paper award on 19th symposium on algorithmic approaches for transportation modelling, optimization, and systems (atmos 2019)'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2019', 'lehilton lelis chaves pedrosa', 'melhor artigo no 4o encontro de teoria da computação - etc'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2019', 'lucas francisco wanner', 'orientador do melhor trabalho de pós-graduação'): ('Melhor trabalho – nível/tipo não informado', 'Premiação / colocação', ''),
    ('2019', 'lucas francisco wanner', 'prêmio excelência didática'): ('Excelência Didática (IC)', 'Premiação / colocação', ''),
    ('2019', 'maria cecília calani baranauskas', 'melhores artigos do ihc2019 (2o. lugar) "tangitime: designing a (socio)enactive experience for deep time in an educational exhibit" (co-autoria de yusseli l.m. mendoza))'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2019', 'maria cecília calani baranauskas', 'melhores artigos no ihc 2019 (mençao honrosa) "reclaiming human space at iot: contributions of the socially aware design" (coautoria de jvsilva)'): ('Best Paper Award / Melhor artigo', 'Menção honrosa', ''),
    ('2019', 'nelson luis saldanha da fonseca', 'google latin america research award'): ('Google Latin America Research Awards (LARA)', 'Premiação / colocação', ''),
    ('2019', 'nelson luis saldanha da fonseca', 'tese doutorado helder m. oliveira - melhor tese doutorado - ctd/sbrc'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2019', 'nelson luis saldanha da fonseca', 'tese doutorado helder m. oliveira - segunda colocada no concurso cltd/clei'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2019', 'paulo lício de geus', 'menção honrosa entre artigos da trilha principal do sbseg 2019, "malware variants identification in practice"'): ('Best Paper Award / Melhor artigo', 'Menção honrosa', ''),
    ('2019', 'pedro jussieu de rezende', '4th place - workshop on open problems and hard instance challenges (em conjunto com n. ramos, r. de jesus, c. de souza, f. usberti)'): ('Outros prêmios e distinções específicos', 'Premiação / colocação', ''),
    ('2019', 'pedro jussieu de rezende', 'segundo lugar no concurso de dissertações de mestrado obtido pelo mestrando natanael ramos (co-orientador)'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2019', 'rafael crivellari saliba schouery', 'best paper award at symposium on algorithmic approaches for transportation modelling, optimization, and systems (atmos)'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2019', 'rafael crivellari saliba schouery', 'orientando de mestrado entre os 10 melhores no xxxii concurso de teses e dissertações da sbc'): ('Dissertações de mestrado', 'Finalista / indicação', ''),
    ('2019', 'ricardo da silva torres', "prêmio de reconhecimento acadêmico 'zeferino vaz'"): ('Prêmio Zeferino Vaz', 'Premiação / colocação', ''),
    ('2019', 'ricardo dahab', 'melhor artigo curto do sbseg 2019'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2019', 'ricardo dahab', 'prêmio destaque da ceseg'): ('Outros prêmios e distinções específicos', 'Premiação / colocação', ''),
    ('2019', 'sandra eliza fontes de avila', 'best paper award, isic skin image analysis workshop, "(de)constructing bias on skin lesion datasets"'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2019', 'sandra eliza fontes de avila', 'google latin america research awards (lara), "improving skin cancer classification with generative adversarial networks - extension"'): ('Google Latin America Research Awards (LARA)', 'Premiação / colocação', ''),
    ('2019', 'sandra eliza fontes de avila', 'prêmio futuro cientista, "detecção de frutos em viticultura utilizando redes neurais e a base de dados embrapa wgisd", aluno: leonardo lima de souza.'): ('Iniciação científica / graduação', 'Premiação / colocação', ''),
    ('2019', 'zanoni dias', 'finalista no concurso de teses e dissertações - categoria mestrado (sbc ctd 2019) (aluno: klairton de lima brito / orientador: zanoni dias)'): ('Dissertações de mestrado', 'Finalista / indicação', ''),
    ('2019', 'zanoni dias', 'prêmio melhor tese de doutorado (aluno: andré rodrigues oliveira / orientador: zanoni dias)'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2020', 'alexandre xavier falcão', 'melhor artigo: badresc: brain anomaly detection based on registration errors and supervoxel classification.'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2020', 'alexandre xavier falcão', 'prêmio de excelência acadêmica'): ('Excelência Acadêmica (IC)', 'Premiação / colocação', ''),
    ('2020', 'alexandre xavier falcão', 'prêmio inventores unicamp 2020 na categoria "tecnologia licenciada", por três tecnologias licenciadas em 2019.'): ('Prêmio Inventores Unicamp', 'Premiação / colocação', ''),
    ('2020', 'anderson de rezende rocha', 'google research awards for latin america ? estudante: antonio theophilo (doutorado)'): ('Google Latin America Research Awards (LARA)', 'Premiação / colocação', ''),
    ('2020', 'anderson de rezende rocha', 'prêmio inventores - patente concedida'): ('Prêmio Inventores Unicamp', 'Premiação / colocação', ''),
    ('2020', 'anderson de rezende rocha', 'top 2% world scientists'): ('Rankings e indicadores de impacto', 'Outro reconhecimento', ''),
    ('2020', 'breno bernard nicolau de frança', 'best paper award - 14th brazilian symposium on software components, architectures, and reuse ? sbcars 2020'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2020', 'breno bernard nicolau de frança', 'noms 2020 best paper award'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2020', 'cecília mary fischer rubira', 'pesquisadora homenageada pela comissão especial de engenharia de software (cees)'): ('Outros prêmios e distinções específicos', 'Outro reconhecimento', ''),
    ('2020', 'edmundo roberto mauro madeira', 'menção honrosa - ctd-doutorado sbrc - ademar takeo akabane; co-orientador: edmundo roberto mauro madeira; orientador: leandro a. villas'): ('Teses de doutorado', 'Menção honrosa', ''),
    ('2020', 'edson borin', 'co-orientador da dissertação de mestrado de antonio carlos guimarães junior. 1o lugar no concurso de teses e dissertações em segurança da informação e sistemas computacionais do xx sbseg'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2020', 'edson borin', 'menção honrosa - 3o melhor artigo do xxi simpósio em sistemas computacionais de alto desempenho. trabalho:'): ('Best Paper Award / Melhor artigo', 'Menção honrosa', ''),
    ('2020', 'edson borin', 'menção honrosa no fórum de pós graduação da erad-sp: trabalho: análise de desempenho dos serviços de armazenamento da nuvem computacional para execução de checkpoint. charles b. rodamilans'): ('Menção honrosa – objeto não especificado', 'Menção honrosa', ''),
    ('2020', 'edson borin', 'orientador da tese de doutorado da flávia pisani. menção honrosa no concurso de teses e dissertações do xviii simpósio em sistemas computacionais de alto desempenho'): ('Teses de doutorado', 'Menção honrosa', ''),
    ('2020', 'edson borin', 'prêmio de reconhecimento docente pela dedicação ao ensino de graduação'): ('Reconhecimento pela Dedicação ao Ensino de Graduação', 'Premiação / colocação', ''),
    ('2020', 'esther luna colombini', '1o lugar no concurso de teses e dissertações (ctd), simpósio brasileiro de robótica (sbr). aluno: gabriel moraes barros, reinforcement and imitation learning applied to autonomous aerial robot control'): ('Teses/dissertações – nível não informado', 'Premiação / colocação', ''),
    ('2020', 'esther luna colombini', '2o lugar no concurso de teses e dissertações (ctd), simpósio brasileiro de robótica (sbr). aluno: hudson martins bruno, lift-slam: um método de slam visual baseado em características com aprendizado p'): ('Teses/dissertações – nível não informado', 'Premiação / colocação', ''),
    ('2020', 'esther luna colombini', 'mérito científico para o trabalho #racegover: automatic image audiodescription'): ('Mérito científico', 'Premiação / colocação', ''),
    ('2020', 'esther luna colombini', 'mérito científico para o trabalho aprendizado de máquina para controle autônomo de trajetória de veículos aéreos não tripulados'): ('Mérito científico', 'Premiação / colocação', ''),
    ('2020', 'fábio luiz usberti', 'medalha de ouro na final brasileira de maratona de programação (atuando como técnico)'): ('Maratona de programação', 'Premiação / colocação', ''),
    ('2020', 'hélio pedrini', '1o lugar - prêmio de iniciação científica (aluno christian massao konishi)'): ('Iniciação científica / graduação', 'Premiação / colocação', ''),
    ('2020', 'hélio pedrini', 'melhor tese de doutorado (aluna: helena de almeida maia)'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2020', 'hélio pedrini', "prêmio de reconhecimento acadêmico 'zeferino vaz'"): ('Prêmio Zeferino Vaz', 'Premiação / colocação', ''),
    ('2020', 'julio cesar dos reis', 'best paper nominee'): ('Best Paper Award / Melhor artigo', 'Finalista / indicação', ''),
    ('2020', 'julio cesar dos reis', 'best paper presentation'): ('Melhor apresentação', 'Premiação / colocação', ''),
    ('2020', 'leandro aparecido villas', 'menção honrosa tese de doutorado no concurso de teses e dissertações do simpósio brasileiro de redes de computadores e sistemas distribuídos'): ('Teses de doutorado', 'Menção honrosa', ''),
    ('2020', 'lucas francisco wanner', 'orientador da melhor dissertação de mestrado - aluno isaias felzmann'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2020', 'nelson luis saldanha da fonseca', 'best paper award, ieee taos tc, o. ciceri, c astrudillo, n. fonseca, dba algorithm for cooperative resource sharing among epon customers'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2020', 'nelson luis saldanha da fonseca', 'google latin america research award'): ('Google Latin America Research Awards (LARA)', 'Premiação / colocação', ''),
    ('2020', 'nelson luis saldanha da fonseca', 'ieee comsoc harold sobol award for exemplery services to meetings and conferences'): ('Outros prêmios e distinções específicos', 'Premiação / colocação', ''),
    ('2020', 'nelson luis saldanha da fonseca', 'tese de felipe forero geometric aspects and random delays in probabilistic broadcasting for wireless ad hoc networks primeira colocada no concurdo clei vi latin american doctoral thesis contest'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2020', 'nelson luis saldanha da fonseca', 'tese oscar ciceri selecionada entre as oito melhores no ctd/sbrc'): ('Teses/dissertações – nível não informado', 'Finalista / indicação', ''),
    ('2020', 'nelson luis saldanha da fonseca', 'tese silvana trindademenção honrosa no ctd/sbrc'): ('Teses/dissertações – nível não informado', 'Menção honrosa', ''),
    ('2020', 'sandra eliza fontes de avila', '1o lugar no concurso de teses e dissertações (ctd), "análise de lesões de pele usando redes generativas adversariais", aluno: alceu bissoto'): ('Teses/dissertações – nível não informado', 'Premiação / colocação', ''),
    ('2020', 'sandra eliza fontes de avila', 'best master?s thesis award, workshop of theses and dissertations (wtd), "improving skin cancer analysis with generative adversarial networks", m.sc.: alceu bissoto'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2020', 'sandra eliza fontes de avila', 'google latin america research awards (lara), "rethinking automated skin cancer classification with unsupervised representation learning ? extension""'): ('Google Latin America Research Awards (LARA)', 'Premiação / colocação', ''),
    ('2020', 'sandra eliza fontes de avila', 'runner-up best paper award, isic skin image analysis workshop, "debiasing skin lesion datasets and models? not so fast"'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2020', 'zanoni dias', 'best paper award'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2021', 'alexandre xavier falcão', 'prêmio de melhor tese de doutorado (aluno: samuel martins)'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2021', 'anderson de rezende rocha', 'google research awards for latin america ? estudante: flavia azevedo (doutorado)'): ('Google Latin America Research Awards (LARA)', 'Premiação / colocação', ''),
    ('2021', 'anderson de rezende rocha', 'google research awards for latin america ? estudante: joão philipe cardenutto (doutorado)'): ('Google Latin America Research Awards (LARA)', 'Premiação / colocação', ''),
    ('2021', 'anderson de rezende rocha', 'prêmio inventores - patente licenciada'): ('Prêmio Inventores Unicamp', 'Premiação / colocação', ''),
    ('2021', 'anderson de rezende rocha', 'top 40 brasilian researchers'): ('Rankings e indicadores de impacto', 'Outro reconhecimento', ''),
    ('2021', 'anderson de rezende rocha', 'top-2 world scientist'): ('Rankings e indicadores de impacto', 'Outro reconhecimento', ''),
    ('2021', 'andré santanchè', 'prêmio menção honrosa artigo curto no sbbd 2021'): ('Best Paper Award / Melhor artigo', 'Menção honrosa', ''),
    ('2021', 'cid carvalho de souza', 'best paper award - liii simpósio brasileiro de pesquisa operacional (sbpo 2021)'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2021', 'claudia maria bauzer medeiros', 'fellow'): ('Fellow / membro / senior member', 'Outro reconhecimento', ''),
    ('2021', 'claudia maria bauzer medeiros', 'mencao honrosa - artigo - addressing search in scientific open data repositories: a semantic metasearch platform (orientando gustavo borges, co-orientador júlio cesar dos reis)'): ('Best Paper Award / Melhor artigo', 'Menção honrosa', ''),
    ('2021', 'claudia maria bauzer medeiros', "prêmio de reconhecimento acadêmico 'zeferino vaz'"): ('Prêmio Zeferino Vaz', 'Premiação / colocação', ''),
    ('2021', 'edmundo roberto mauro madeira', 'best paper award - sixth international workshop on process querying, manipulation, and intelligence (pqmi 2021)'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2021', 'edson borin', 'best paper award - clap-bot: a framework for automatic optimization of high-performance elastic applications on the clouds'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2021', 'edson borin', 'best student paper award - leveraging vcpu-utilization rates to select cost-efficient vms for parallel workloads. student: william f. c. tavares'): ('Best Student Paper Award', 'Premiação / colocação', ''),
    ('2021', 'edson borin', 'menção honrosa. artigo: fast and low-cost search for efficient cloud configurations for hpc workloads'): ('Best Paper Award / Melhor artigo', 'Menção honrosa', ''),
    ('2021', 'fábio luiz usberti', 'medalha de bronze na final brasileira de maratona de programação (atuando como técnico)'): ('Maratona de programação', 'Premiação / colocação', ''),
    ('2021', 'fábio luiz usberti', 'menção honrosa no concurso de prêmio de dissertação de mestrado (como orientador)'): ('Dissertações de mestrado', 'Menção honrosa', ''),
    ('2021', 'flávio keidi miyazawa', 'orientador da dissertação ganhadora de menção honrosa no concurso de teses e dissertações da sbc, matheus j. ota'): ('Dissertações de mestrado', 'Menção honrosa', ''),
    ('2021', 'flávio keidi miyazawa', 'orientador da dissertação ganhadora do prêmio de melhor dissertação do ic-unicamp, matheus j. ota'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2021', 'flávio keidi miyazawa', 'orientador da dissertação ganhadora do prêmio de melhor dissertação do liii simpósio brasileiro de pesquisa operacional, matheus j. ota'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2021', 'guido costa souza de araújo', 'prêmio de excelência didática'): ('Excelência Didática (IC)', 'Premiação / colocação', ''),
    ('2021', 'guido costa souza de araújo', 'prêmio de reconhecimento docente pela dedicação ao ensino de graduação'): ('Reconhecimento pela Dedicação ao Ensino de Graduação', 'Premiação / colocação', ''),
    ('2021', 'leandro aparecido villas', 'dissertação de mestrado selecionada entre as 12 melhores no concurso de teses e dissertações do simpósio brasileiro de redes de computadores e sistemas distribuídos'): ('Dissertações de mestrado', 'Finalista / indicação', ''),
    ('2021', 'leandro aparecido villas', 'prêmio de melhor tese de doutorado no concurso de teses e dissertações do simpósio brasileiro de redes de computadores e sistemas distribuídos'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2021', 'leandro aparecido villas', 'tese de doutorado selecionada entre as seis melhores no concurso de teses e dissertações do simpósio brasileiro de redes de computadores e sistemas distribuídos'): ('Teses de doutorado', 'Finalista / indicação', ''),
    ('2021', 'lucas francisco wanner', 'best interactive presentation'): ('Melhor apresentação', 'Premiação / colocação', ''),
    ('2021', 'maria cecília calani baranauskas', 'acm chi social impact award'): ('Outros prêmios e distinções específicos', 'Premiação / colocação', ''),
    ('2021', 'nelson luis saldanha da fonseca', 'menção honrosa no wgrs/sbrc alex santos, denis rosário, eduardo cerqueira, nelson fonseca, "a strategy to the reduction of communication overhead and overfitting in federated learning"'): ('Menção honrosa – objeto não especificado', 'Menção honrosa', ''),
    ('2021', 'nelson luis saldanha da fonseca', 'tese de mestrado de daniela maria casas velasco selecionada como finalista no ctd/sbrc'): ('Dissertações de mestrado', 'Finalista / indicação', ''),
    ('2021', 'nelson luis saldanha da fonseca', 'tese de mestrado de william f. villota jácome selecionada como finalista no ctd/sbrc'): ('Dissertações de mestrado', 'Finalista / indicação', ''),
    ('2021', 'ricardo de oliveira anido', 'prêmio proec de extensão universitária'): ('Prêmio ProEC de Extensão Universitária', 'Premiação / colocação', ''),
    ('2021', 'rodolfo jardim de azevedo', 'best paper award'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2021', 'sandra eliza fontes de avila', 'bolsa de produtividade em pesquisa do cnpq no nível 2'): ('Bolsa de produtividade CNPq', 'Outro reconhecimento', ''),
    ('2021', 'sandra eliza fontes de avila', 'google latin america research awards (lara), "the quest for generalization in small medical datasets""'): ('Google Latin America Research Awards (LARA)', 'Premiação / colocação', ''),
    ('2021', 'zanoni dias', 'best paper award'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2021', 'zanoni dias', 'prêmio melhor dissertação de mestrado (aluno: gabriel bianchin de oliveira / orientador: zanoni dias)'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2022', 'alexandre xavier falcão', 'prêmio de melhor dissertação de mestrado (aluno: jordão bragantini)'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2022', 'alexandre xavier falcão', 'prêmio inventores unicamp 2022 - tecnologia licenciada'): ('Prêmio Inventores Unicamp', 'Premiação / colocação', ''),
    ('2022', 'alexandre xavier falcão', 'senior ieee member'): ('Fellow / membro / senior member', 'Outro reconhecimento', ''),
    ('2022', 'anderson de rezende rocha', '2022 master thesis, prêmio cepetro em pesquisa, desenvolvimento e inovação'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2022', 'breno bernard nicolau de frança', '2nd best paper award (sbcars)'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2022', 'breno bernard nicolau de frança', 'distinguished pc member award (sbcars)'): ('Revisor ou membro de comitê de destaque', 'Premiação / colocação', ''),
    ('2022', 'cid carvalho de souza', 'terceiro lugar no xxxv concurso de teses e dissertações (ctd) do xlii congresso da sbc obtido pelo mestrando felipe de c. pereira (coorientador)'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2022', 'edson borin', 'orientador da tese de doutorado de vanderson martins do rosário. 1o lugar no concurso de teses e dissertações do xxiii simpósio em sistemas computacionais de alto desempenho'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2022', 'esther luna colombini', '1o lugar no concurso de teses e dissertações (ctd), simpósio brasileiro de robótica (sbr). aluno: leonardo de lellis rossi, aprendizado procedimental e sensório-motor em robôs cognitivos'): ('Teses/dissertações – nível não informado', 'Premiação / colocação', ''),
    ('2022', 'esther luna colombini', '2o lugar no concurso de teses e dissertações (ctd), simpósio brasileiro de robótica (sbr). aluno: patrick de carvalho, odometria multifrequência para veículos aéreos autônomos'): ('Teses/dissertações – nível não informado', 'Premiação / colocação', ''),
    ('2022', 'esther luna colombini', '3rd best paper - xix ieee latin american robotics symposium - memory-based deep reinforcement learning for humanoid locomotion under noisy scenarios'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2022', 'esther luna colombini', 'prêmio proec de extensão universitária'): ('Prêmio ProEC de Extensão Universitária', 'Premiação / colocação', ''),
    ('2022', 'fábio luiz usberti', 'duas medalhas de ouro na final brasileira de maratona de programação (atuando como técnico)'): ('Maratona de programação', 'Premiação / colocação', ''),
    ('2022', 'fábio luiz usberti', 'livre docência em complexidade de algoritmos'): ('Outros prêmios e distinções específicos', 'Outro reconhecimento', ''),
    ('2022', 'fábio luiz usberti', 'menção honrosa por ter sido por 5 anos técnico de times na final mundial de maratona de programação'): ('Maratona de programação', 'Menção honrosa', ''),
    ('2022', 'flávio keidi miyazawa', 'orientador da tese ganhadora de menção honrosa no concurso de teses da sbmac, vinícius loti de lima'): ('Teses de doutorado', 'Menção honrosa', 'Nível inferido dos outros registros de 2022 do mesmo aluno Vinícius Loti de Lima.'),
    ('2022', 'flávio keidi miyazawa', 'orientador do 1o. colocado no concurso de teses de doutorado da sbc, vinícius loti de lima'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2022', 'flávio keidi miyazawa', 'orientador do 1o. colocado no concurso de teses de doutorado da sbpo, vinícius loti de lima'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2022', 'guido costa souza de araújo', "prêmio de reconhecimento acadêmico 'zeferino vaz'"): ('Prêmio Zeferino Vaz', 'Premiação / colocação', ''),
    ('2022', 'hervé cédric yviquel', 'best poster award, latin american high performance computing conference (carla 2022), an openmp-only linear algebra library for distributed architectures'): ('Best Poster Award / Melhor pôster', 'Premiação / colocação', ''),
    ('2022', 'julio cesar dos reis', '3a melhor dissertação (orientando andré regino)'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2022', 'julio cesar dos reis', 'best paper award'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2022', 'julio cesar dos reis', 'best student paper award'): ('Best Student Paper Award', 'Premiação / colocação', ''),
    ('2022', 'leandro aparecido villas', 'bolsa de produtividade em pesquisa do cnpq no nível 1d'): ('Bolsa de produtividade CNPq', 'Outro reconhecimento', ''),
    ('2022', 'leandro aparecido villas', 'menção honrosa no prêmio capes de tese'): ('Teses de doutorado', 'Menção honrosa', ''),
    ('2022', 'leandro aparecido villas', 'prêmio de melhor artigo no international conference on enterprise information systems'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2022', 'leandro aparecido villas', 'prêmio de melhor tese de doutorado do instituto de computação da unicamp'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2022', 'leandro aparecido villas', 'prêmio de melhor tese de doutorado entre os institutos de computação das universidades de berna, friburgo e neuchate'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2022', 'leandro aparecido villas', 'prêmio de melhor tese de doutorado no concurso de teses e dissertações na latin american computing conference'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2022', 'leandro aparecido villas', 'segundo lugar no concurso de teses e dissertações do congresso da sociedade brasileira da computação'): ('Teses/dissertações – nível não informado', 'Premiação / colocação', ''),
    ('2022', 'lucas francisco wanner', 'prêmio de reconhecimento docente pela dedicação ao ensino de graduação'): ('Reconhecimento pela Dedicação ao Ensino de Graduação', 'Premiação / colocação', ''),
    ('2022', 'luiz fernando bittencourt', 'best paper award - courb workshop'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2022', 'luiz fernando bittencourt', 'best reviewer award noms 2022'): ('Revisor ou membro de comitê de destaque', 'Premiação / colocação', ''),
    ('2022', 'maria cecília calani baranauskas', 'melhor artigo na trilha 1: relatos de experiências educacionais usando tdic" do xxviii workshop de informática na escola- wie 2022.https://doi.org/10.5753/wie.2022.224733'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2022', 'maria cecília calani baranauskas', 'melhor artigo na trilha 3: fatores humanos em tecnologia digital na educação, do xxxiii simpósio brasileiro de informática na educação - sbie 2022. https://doi.org/10.5753/sbie.2022.225675'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2022', 'nelson luis saldanha da fonseca', 'excelência acadêmica'): ('Excelência Acadêmica (IC)', 'Premiação / colocação', ''),
    ('2022', 'nelson luis saldanha da fonseca', 'tese carlos alberto astudillo trujillo segundo colocado no vi concurso latino-americano de tese de doutorado'): ('Teses de doutorado', 'Premiação / colocação', 'Edição VI indicada na fonte, diferente da edição VIII no outro registro; mantido.'),
    ('2022', 'nelson luis saldanha da fonseca', 'tese de doutorado de carlos alberto astudillo trujillo - segundo lugar no viii concurso latinoamericano de tesis de doctorado'): ('Teses de doutorado', 'Premiação / colocação', 'Edição VIII indicada na fonte, diferente da edição VI no outro registro; mantido.'),
    ('2022', 'pedro jussieu de rezende', 'terceiro lugar no xxxv concurso de teses e dissertações (ctd) do xlii congresso da sociedade brasileira de computação obtido pelo mestrando felipe de c. pereira (orientador)'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2022', 'rodolfo jardim de azevedo', 'melhor artigo'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2022', 'sandra eliza fontes de avila', 'google award for inclusion research (air), "dark skin matters: fair and unbiased skin lesion models"'): ('Google Award for Inclusion Research (AIR)', 'Premiação / colocação', ''),
    ('2022', 'sandra eliza fontes de avila', 'runner-up best paper award, isic skin image analysis workshop, "an evaluation of self-supervised pre-training for skin-lesion analysis?'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2022', 'sandra eliza fontes de avila', 'top 2% of most-cited researchers in the world'): ('Rankings e indicadores de impacto', 'Outro reconhecimento', ''),
    ('2022', 'zanoni dias', 'finalista no concurso de teses e dissertações - categoria mestrado (sbc ctd 2022) (aluno: gabriel henriques siqueira / orientador: zanoni dias)'): ('Dissertações de mestrado', 'Finalista / indicação', ''),
    ('2023', 'allan mariano de souza', 'best paper award xli simpósio brasileiro de redes de computadores e sistemas distribuídos (sbrc)'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2023', 'anderson de rezende rocha', 'asia pacific artificial intelligence fellow'): ('Fellow / membro / senior member', 'Outro reconhecimento', ''),
    ('2023', 'anderson de rezende rocha', 'ieee fellow'): ('Fellow / membro / senior member', 'Outro reconhecimento', ''),
    ('2023', 'anderson de rezende rocha', 'vinfuture prize nomination counselor'): ('Eleição / nomeação para cargo ou comitê', 'Outro reconhecimento', ''),
    ('2023', 'andré santanchè', 'prêmio de melhor pesquisa/pôster na modalidade de tecnologias no conicc'): ('Best Poster Award / Melhor pôster', 'Premiação / colocação', ''),
    ('2023', 'bruno barbieri de pontes cafeo', '1o lugar - mestrado - concurso de teses e dissertações em engenharia de software (ctd-es) do xiv cbsoft 2023 - candidato: gabriel santana de menezes. orientador: bruno cafeo, co-orientador: andré hora'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2023', 'carlos alberto astudillo trujillo', 'primeiro lugar na sexta edição do concurso de teses e dissertações (ctd) do simposio brasileiro de redes de computadores e sistemas distribuídos (sbrc 2023)'): ('Teses de doutorado', 'Premiação / colocação', 'Nível inferido do registro de Nelson Fonseca referente ao mesmo concurso e ao mesmo premiado.'),
    ('2023', 'cid carvalho de souza', "prêmio de reconhecimento acadêmico 'zeferino vaz'"): ('Prêmio Zeferino Vaz', 'Premiação / colocação', ''),
    ('2023', 'edmundo roberto mauro madeira', 'excelência acadêmica'): ('Excelência Acadêmica (IC)', 'Premiação / colocação', ''),
    ('2023', 'emanuel felipe duarte', 'artigo entre os 25 artigos mais citados dos ihcs em 25 anos de realizações do evento'): ('Rankings e indicadores de impacto', 'Outro reconhecimento', ''),
    ('2023', 'esther luna colombini', 'prêmio de reconhecimento docente pela dedicação ao ensino de graduação'): ('Reconhecimento pela Dedicação ao Ensino de Graduação', 'Premiação / colocação', ''),
    ('2023', 'fábio luiz usberti', '2o lugar entre os melhores trabalhos do li sbpo'): ('Melhor trabalho – nível/tipo não informado', 'Premiação / colocação', ''),
    ('2023', 'fábio luiz usberti', 'uma medalha de ouro e uma medalha de bronze na final brasileira de maratona de programação 2023 (atuando como técnico)'): ('Maratona de programação', 'Premiação / colocação', ''),
    ('2023', 'flávio keidi miyazawa', 'orientador de tese selecionada entre as 5 melhores no concurso de teses da sbpo, mauro h. mulati'): ('Teses/dissertações – nível não informado', 'Finalista / indicação', ''),
    ('2023', 'hervé cédric yviquel', 'best paper award, waccpd 23, memory transfer decomposition: exploring smart data movement through architecture-aware strategies'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2023', 'hervé cédric yviquel', 'honorable mention, erad 23, simplifying hpc application development with openmp cluster and unified memory'): ('Menção honrosa – objeto não especificado', 'Menção honrosa', ''),
    ('2023', 'juliana freitag borin', 'dissertação de mestrado de maria julia berriel de sousa menção honrosa no concurso ctd/csbc'): ('Dissertações de mestrado', 'Menção honrosa', 'Consolidado com registro de finalista do mesmo concurso, aluna e ano; preservada a menção honrosa.'),
    ('2023', 'juliana freitag borin', 'dissertação de mestrado de maria julia berriel de sousa selecionada entre as 10 melhores do país no concurso ctd/csbc'): ('Dissertações de mestrado', 'Finalista / indicação', ''),
    ('2023', 'julio cesar dos reis', 'best paper award'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2023', 'julio cesar dos reis', 'best paper candidate'): ('Best Paper Award / Melhor artigo', 'Finalista / indicação', ''),
    ('2023', 'julio cesar dos reis', 'best student paper award nominee'): ('Best Student Paper Award', 'Finalista / indicação', ''),
    ('2023', 'julio césar lópez hernández', 'menção honrosa no concurso: prêmio tese destaque unicamp 2022. tese orientada.'): ('Teses/dissertações – nível não informado', 'Menção honrosa', ''),
    ('2023', 'julio césar lópez hernández', 'prêmio de melhor tese de doutorado orientada'): ('Teses de doutorado', 'Premiação / colocação', 'Possível duplicidade com descrição de tese orientada de 2022; evento e aluno não informados.'),
    ('2023', 'julio césar lópez hernández', 'prêmio de melhor tese de doutorado orientada-2022'): ('Teses de doutorado', 'Premiação / colocação', 'Possível duplicidade com outra descrição de tese orientada; evento e aluno não informados.'),
    ('2023', 'leandro aparecido villas', 'best paper award na 25th international conference on human-computer interaction'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2023', 'luiz fernando bittencourt', 'best paper award - xli simpósio brasileiro de redes de computadores e sistemas distribuídos (sbrc 2023)'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2023', 'nelson luis saldanha da fonseca', 'artigo chantre, hernani d.; da fonseca, nelson l. s, "design of 5g mec-based networks with 1:n:k protection scheme" runner-up informstechnical section on telecommunications and network analytics'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', 'Consolidado com a descrição do mesmo artigo e evento, que não especifica a colocação.'),
    ('2023', 'nelson luis saldanha da fonseca', 'artigo chantre, hernani d.; da fonseca, nelson l. s,, "design of 5g mec-based networks with 1:n:k protection - technical section on telecommunications and network analytics best paper award 2022/'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2023', 'nelson luis saldanha da fonseca', 'ieee communications society education award'): ('Outros prêmios e distinções específicos', 'Premiação / colocação', ''),
    ('2023', 'nelson luis saldanha da fonseca', 'tese de doutorado de carlos alberto astudillo trujillo - primeiro lugar no sexto concurso de teses de doutorado (ctd) do sbrc 2023'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2023', 'nelson luis saldanha da fonseca', 'tese de doutorado de rodrigo augusto cardoso da silva - terceiro lugar no sexto concurso de teses de doutorado (ctd) do sbrc 2023'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2023', 'nelson luis saldanha da fonseca', 'tese de doutorado rodrigo augusto cardoso da silva - primeiro lugar ctd ix concurso latinoamericano de tesis de doctorado da xlix conferencia latinoamericana de informatica'): ('Teses de doutorado', 'Premiação / colocação', 'Consolidado com outra descrição do mesmo aluno, concurso IX e primeiro lugar.'),
    ('2023', 'nelson luis saldanha da fonseca', 'tese doutorado rodrigo augusto cardoso da silva primeiro lugar no cltd - ix concurso latinoamericano de de tesis de doctorado'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2023', 'paulo lício de geus', 'melhor tese de doutorado de 2022, ufpr-inf'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2023', 'paulo lício de geus', 'prêmio proec de extensão universitária'): ('Prêmio ProEC de Extensão Universitária', 'Premiação / colocação', ''),
    ('2023', 'rafael crivellari saliba schouery', 'melhor artigo no viii encontro de teoria da computação'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2023', 'rafael crivellari saliba schouery', 'prêmio excelência didática'): ('Excelência Didática (IC)', 'Premiação / colocação', ''),
    ('2023', 'rubén interian kovaliova', 'chosen for the special virtual 30th anniversary issue as one of 30 representative and influential papers in the journal\'s history, "a decision support system for fraud detection in public procurement"'): ('Rankings e indicadores de impacto', 'Outro reconhecimento', ''),
    ('2023', 'sandra eliza fontes de avila', 'prêmio de excelência didática'): ('Excelência Didática (IC)', 'Premiação / colocação', ''),
    ('2023', 'sandra eliza fontes de avila', 'reconhecimento acadêmico em direitos humanos unicamp-instituto vladimir herzog, "leveraging self-supervised learning for scene recognition in child sexual abuse imagery"'): ('Reconhecimento em Direitos Humanos', 'Premiação / colocação', ''),
    ('2023', 'sandra eliza fontes de avila', 'top 2% of most-cited researchers in the world'): ('Rankings e indicadores de impacto', 'Outro reconhecimento', ''),
    ('2024', 'alexandre xavier falcão', 'menção honrosa no workshop de teses e dissertações do sibgrapi 2024 com a tese de doutorado de b. benato. pseudo labeling and classification of high-dimensional data using visual analytics'): ('Teses de doutorado', 'Menção honrosa', ''),
    ('2024', 'alexandre xavier falcão', 'prêmio de melhor artigo da trilha principal (sibgrapi 2024), a. nobrega et al., improving local latent fingerprint representations under data constraints'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2024', 'alexandre xavier falcão', 'prêmio inventor 2024 na categoria tecnologia absorvida pelo mercado com 4 tecnologias'): ('Prêmio Inventores Unicamp', 'Premiação / colocação', ''),
    ('2024', 'allan mariano de souza', 'best paper viii workshop de computação urbana'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2024', 'anderson de rezende rocha', 'prêmio proec de extensão universitária'): ('Prêmio ProEC de Extensão Universitária', 'Premiação / colocação', ''),
    ('2024', 'breno bernard nicolau de frança', 'distinguished reviewer (sbcars)'): ('Revisor ou membro de comitê de destaque', 'Premiação / colocação', ''),
    ('2024', 'edson borin', 'medalha da ordem do mérito de egresso'): ('Outros prêmios e distinções específicos', 'Premiação / colocação', ''),
    ('2024', 'edson borin', 'orientador da tese de doutorado de antonio carlos guimarães jr.. 1o lugar no concurso de teses e dissertações do xxv simpósio em sistemas computacionais de alto desempenho'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2024', 'emanuel felipe duarte', 'revisor destaque no xxiii simpósio brasileiro sobre fatores humanos em sistemas computacionais (ihc 2024)'): ('Revisor ou membro de comitê de destaque', 'Premiação / colocação', ''),
    ('2024', 'esther luna colombini', '2nd best paper award para o artigo i choose you, reinforcement learning! trained rl agents for pokémon battles no sbgames'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2024', 'esther luna colombini', '3rd best paper award para o artigo improving task-incremental human activity recognition with plasticity techniques no kdmile/bracis'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2024', 'fábio luiz usberti', 'medalha de bronze na final brasileira de maratona de programação 2024 (atuando como técnico)'): ('Maratona de programação', 'Premiação / colocação', ''),
    ('2024', 'fábio luiz usberti', 'prêmio de reconhecimento docente pela dedicação ao ensino de graduação'): ('Reconhecimento pela Dedicação ao Ensino de Graduação', 'Premiação / colocação', ''),
    ('2024', 'hélio pedrini', 'prêmio - excelência acadêmica do instituto de computação'): ('Excelência Acadêmica (IC)', 'Premiação / colocação', ''),
    ('2024', 'hélio pedrini', 'prêmio inventores - tecnologias licenciadas'): ('Prêmio Inventores Unicamp', 'Premiação / colocação', ''),
    ('2024', 'julio cesar dos reis', 'best paper award'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2024', 'julio cesar dos reis', 'best paper contender'): ('Best Paper Award / Melhor artigo', 'Finalista / indicação', ''),
    ('2024', 'julio césar lópez hernández', 'menção honrosa no concurso: ctdseg-2024. tese orientada.'): ('Teses/dissertações – nível não informado', 'Menção honrosa', ''),
    ('2024', 'leandro aparecido villas', 'melhor artigo no viii workshop de computação urbana do xlii simpósio brasileiro de redes de computadores e sistemas distribuídos'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2024', 'leandro aparecido villas', 'menção honrosa no prêmio tese destaque unicamp. ciências exatas e da terra.'): ('Teses/dissertações – nível não informado', 'Menção honrosa', ''),
    ('2024', 'leandro aparecido villas', 'menção honrosa no workshop de trabalhos de iniciação científica e de graduação do xlii simpósio brasileiro de redes de computadores e sistemas distribuídos'): ('Iniciação científica / graduação', 'Menção honrosa', ''),
    ('2024', 'leandro aparecido villas', 'prêmio de melhor tese no concurso de teses e dissertações do instituto de computação da unicamp'): ('Teses/dissertações – nível não informado', 'Premiação / colocação', ''),
    ('2024', 'leandro aparecido villas', 'segundo melhor artigo no viii workshop de computação urbana do xlii simpósio brasileiro de redes de computadores e sistemas distribuídos de 2024'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2024', 'leandro aparecido villas', 'terceiro lugar no concurso de trabalhos de iniciação científica do congresso da sociedade brasileira de computação 2024'): ('Iniciação científica / graduação', 'Premiação / colocação', ''),
    ('2024', 'luiz fernando bittencourt', 'lucas de paula soares - melhor iniciação científica do instituto de computação da unicamp'): ('Iniciação científica / graduação', 'Premiação / colocação', ''),
    ('2024', 'luiz fernando bittencourt', 'prêmio de reconhecimento acadêmico zeferino vaz'): ('Prêmio Zeferino Vaz', 'Premiação / colocação', ''),
    ('2024', 'nelson luis saldanha da fonseca', 'elsevier standardized citation indicators of the 100,000 scientists by c-score in all sub-fields or a percentile rank of 2% or above in the sub-field'): ('Rankings e indicadores de impacto', 'Outro reconhecimento', ''),
    ('2024', 'nelson luis saldanha da fonseca', 'tese de hernani delgado chantre ?mec placement problem in protected 5g networks? terceira colocada no concurso clei x latin american doctoral thesis contest, 2024.'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2024', 'rafael crivellari saliba schouery', 'orientando entre os 10 finalistas do concurso de trabalhos de iniciação científica do congresso da sociedade brasileira de computação'): ('Iniciação científica / graduação', 'Finalista / indicação', ''),
    ('2024', 'rafael crivellari saliba schouery', 'prêmio de iniciação científica do simpósio brasileiro de pesquisa operacional'): ('Iniciação científica / graduação', 'Premiação / colocação', ''),
    ('2024', 'sandra eliza fontes de avila', 'reconhecimento acadêmico em direitos humanos unicamp-instituto vladimir herzog, "transformers-based few-shot learning for scene classification in child sexual abuse imagery"'): ('Reconhecimento em Direitos Humanos', 'Premiação / colocação', ''),
    ('2024', 'santiago valdés ravelo', 'excelência didática'): ('Excelência Didática (IC)', 'Premiação / colocação', ''),
    ('2025', 'alexandre xavier falcão', 'menção honrosa no workshop de patologia digital e computacional (sibgrapi 2025) com o trabalho: fabio kochem et al. domain-specific seed removal for biomedical image segmentation using sicl'): ('Menção honrosa – objeto não especificado', 'Menção honrosa', ''),
    ('2025', 'alexandre xavier falcão', 'menção honrosa no workshop de teses e dissertações do sibgrapi 2025 com o trabalho de doutorado. felipe de c. belem. towards an agnostic superspel segmentation framework'): ('Teses de doutorado', 'Menção honrosa', ''),
    ('2025', 'alexandre xavier falcão', 'prêmio inventores 2025 por depósito de patente'): ('Prêmio Inventores Unicamp', 'Premiação / colocação', ''),
    ('2025', 'allan mariano de souza', 'menção honrosa - xliii simpósio brasileiro de redes de computadores e sistemas distribuídos (sbrc 2025)'): ('Menção honrosa – objeto não especificado', 'Menção honrosa', ''),
    ('2025', 'anderson de rezende rocha', 'agentic ai innovation awards'): ('Outros prêmios e distinções específicos', 'Premiação / colocação', ''),
    ('2025', 'anderson de rezende rocha', 'chair (2025-2026) - ieee information forensics and security technical committee'): ('Eleição / nomeação para cargo ou comitê', 'Outro reconhecimento', ''),
    ('2025', 'anderson de rezende rocha', 'ieee biometrics council distinguished lecturer'): ('Keynote / distinguished lecturer', 'Outro reconhecimento', ''),
    ('2025', 'anderson de rezende rocha', 'melhor tese de doutorado ? estudante: gabriel bertocco'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2025', 'anderson de rezende rocha', 'national academy of artificial intelligence elected member'): ('Fellow / membro / senior member', 'Outro reconhecimento', ''),
    ('2025', 'anderson de rezende rocha', 'top 10 professionals in technology and innovation'): ('Rankings e indicadores de impacto', 'Outro reconhecimento', ''),
    ('2025', 'anderson de rezende rocha', 'top 2% world scientists'): ('Rankings e indicadores de impacto', 'Outro reconhecimento', ''),
    ('2025', 'anderson de rezende rocha', 'unicamp inventor award ? licensed technology. unicamp inventor award for the software: viva sensing'): ('Prêmio Inventores Unicamp', 'Premiação / colocação', ''),
    ('2025', 'anderson de rezende rocha', 'zeferino vaz unicamp award'): ('Prêmio Zeferino Vaz', 'Premiação / colocação', ''),
    ('2025', 'andré santanchè', 'prêmio menção honroza no bsb 2025'): ('Menção honrosa – objeto não especificado', 'Menção honrosa', ''),
    ('2025', 'andré santanchè', 'prêmio proec de extensão universitária'): ('Prêmio ProEC de Extensão Universitária', 'Premiação / colocação', ''),
    ('2025', 'breno bernard nicolau de frança', 'prêmio inventores unicamp - categoria: propriedade intelectual licenciada'): ('Prêmio Inventores Unicamp', 'Premiação / colocação', ''),
    ('2025', 'breno bernard nicolau de frança', 'revisor de destaque (sbcars)'): ('Revisor ou membro de comitê de destaque', 'Premiação / colocação', ''),
    ('2025', 'breno bernard nicolau de frança', 'revisor de destaque (sbes-iier)'): ('Revisor ou membro de comitê de destaque', 'Premiação / colocação', ''),
    ('2025', 'bruno barbieri de pontes cafeo', 'distinguished reviewer'): ('Revisor ou membro de comitê de destaque', 'Premiação / colocação', ''),
    ('2025', 'carlos alberto astudillo trujillo', 'melhor artigo do xxx workshop de gerência e operação de redes e serviços (wgrs 2025)'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2025', 'claudia maria bauzer medeiros', 'keynote speaker - international conference process mining'): ('Keynote / distinguished lecturer', 'Outro reconhecimento', ''),
    ('2025', 'emanuel felipe duarte', 'menção honrosa da trilha cultura, na categoria artigo completo, no xxiv simpósio brasileiro de jogos e entretenimento digital (sbgames 2025)'): ('Best Paper Award / Melhor artigo', 'Menção honrosa', ''),
    ('2025', 'esther luna colombini', '1o lugar - melhor tese de doutorado - no concurso de teses e dissertações (ctdr) da sbc, aluna: letícia mara berto,'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2025', 'esther luna colombini', '2o lugar - melhor tese de doutorado - no concurso de teses e dissertações (ctdr) da sbc, aluna: alana de santana correia,'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2025', 'fábio luiz usberti', 'medalha de ouro na final brasileira de maratona de programação 2024 (atuando como técnico)'): ('Maratona de programação', 'Premiação / colocação', 'Ano da tabela 2025; descrição menciona evento de 2024.'),
    ('2025', 'flávio keidi miyazawa', 'orientador do 3o. colocado no concurso latinoamericano de teses de mestrado do clei, lucas guesser, cltm-clei.'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2025', 'hélio pedrini', 'melhor dissertação de mestrado (aluna: jhéssica victória santos da silva)'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2025', 'hélio pedrini', 'prêmio - reconhecimento docente pela dedicação ao ensino de graduação'): ('Reconhecimento pela Dedicação ao Ensino de Graduação', 'Premiação / colocação', ''),
    ('2025', 'hélio pedrini', 'prêmio jabuti: livro indicado como um dos 10 finalistas na categoria ciência da computação'): ('Outros prêmios e distinções específicos', 'Finalista / indicação', ''),
    ('2025', 'hilder vitor lima pereira', 'best short paper in cryptography at sbseg 2025'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2025', 'isaías bittencourt felzmann', 'melhor artigo'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2025', 'islene calciolari garcia', 'menção honrosa por publicação de artigo'): ('Best Paper Award / Melhor artigo', 'Menção honrosa', ''),
    ('2025', 'juliana freitag borin', 'melhor artigo da trilha relatos de experiência do xix wit - women in information technology'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2025', 'juliana freitag borin', 'menção honrosa no ix worshop de computação urbana do simpósio brasileiro de redes de computadores e sistemas distribuídos'): ('Menção honrosa – objeto não especificado', 'Menção honrosa', ''),
    ('2025', 'julio cesar dos reis', 'best paper award'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2025', 'julio cesar dos reis', 'menção honrosa - doutorado (orientando andré gomes regino)'): ('Teses de doutorado', 'Menção honrosa', ''),
    ('2025', 'julio césar lópez hernández', "melhor artigo completo da trilha de criptografia do xxv simpósio brasileiro de cibersegurança (sbseg'25)"): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2025', 'leandro aparecido villas', 'segundo melhor artigo na trilha principal do simpósio brasileiro de redes de computadores e sistemas distribuídos de 2025, sbc.'): ('Best Paper Award / Melhor artigo', 'Premiação / colocação', ''),
    ('2025', 'nelson luis saldanha da fonseca', 'elsevier standardized citation indicators of the 100,000 scientists by c-score in all sub-fields or a percentile rank of 2% or above in the sub-field'): ('Rankings e indicadores de impacto', 'Outro reconhecimento', ''),
    ('2025', 'nelson luis saldanha da fonseca', 'mérito científico'): ('Mérito científico', 'Premiação / colocação', ''),
    ('2025', 'sandra eliza fontes de avila', 'best poster award, ?back to the basics on predicting transfer performance?, levy gurgel chaves'): ('Best Poster Award / Melhor pôster', 'Premiação / colocação', ''),
    ('2025', 'sandra eliza fontes de avila', 'dermatologia + inclusiva, "peles negras importam: modelos de inteligência artificial para análise de lesões de pele negra"'): ('Outros prêmios e distinções específicos', 'Premiação / colocação', ''),
    ('2025', 'sandra eliza fontes de avila', 'prêmio de melhor dissertação de mestrado, ?avaliação de ferramentas de ética em modelos de linguagem desenvolvidos em português?, jhessica victoria santos da silva'): ('Dissertações de mestrado', 'Premiação / colocação', ''),
    ('2025', 'zanoni dias', 'prêmio melhor tese de doutorado (aluno: alexsandro oliveira alexandrino / orientador: zanoni dias)'): ('Teses de doutorado', 'Premiação / colocação', ''),
    ('2025', 'zanoni dias', 'terceiro lugar no concurso de teses e dissertações - categoria doutorado (sbc ctd 2025) (aluno: alexsandro oliveira alexandrino / orientador: zanoni dias)'): ('Teses de doutorado', 'Premiação / colocação', ''),
}

# Descrições repetidas do mesmo reconhecimento: chave excluída -> chave mantida.
DUPLICADOS = {
    ('2023', 'nelson luis saldanha da fonseca', 'artigo chantre, hernani d.; da fonseca, nelson l. s,, "design of 5g mec-based networks with 1:n:k protection - technical section on telecommunications and network analytics best paper award 2022/'): ('2023', 'nelson luis saldanha da fonseca', 'artigo chantre, hernani d.; da fonseca, nelson l. s, "design of 5g mec-based networks with 1:n:k protection scheme" runner-up informstechnical section on telecommunications and network analytics'),
    ('2023', 'nelson luis saldanha da fonseca', 'tese doutorado rodrigo augusto cardoso da silva primeiro lugar no cltd - ix concurso latinoamericano de de tesis de doctorado'): ('2023', 'nelson luis saldanha da fonseca', 'tese de doutorado rodrigo augusto cardoso da silva - primeiro lugar ctd ix concurso latinoamericano de tesis de doctorado da xlix conferencia latinoamericana de informatica'),
    ('2023', 'juliana freitag borin', 'dissertação de mestrado de maria julia berriel de sousa selecionada entre as 10 melhores do país no concurso ctd/csbc'): ('2023', 'juliana freitag borin', 'dissertação de mestrado de maria julia berriel de sousa menção honrosa no concurso ctd/csbc'),
}


def gravar(path, linhas, colunas):
    with path.open('w', encoding='utf-8-sig', newline='') as arq:
        w = csv.DictWriter(arq, fieldnames=colunas, delimiter=';', extrasaction='ignore')
        w.writeheader()
        w.writerows(linhas)


def classificar_premios(entrada, saida):
    """Classifica o CSV consolidado, elimina duplicidades adicionais e gera relatórios."""
    saida.mkdir(parents=True, exist_ok=True)
    with entrada.open(encoding='utf-8-sig', newline='') as arq:
        reader = csv.DictReader(arq, delimiter=';')
        campos = reader.fieldnames
        linhas = list(reader)
    chaves = {chave(r) for r in linhas}
    vistos = set()
    detalhado, excluidos = [], []
    for r in linhas:
        k = chave(r)
        c, s, nota = REGRAS.get(k, ('Revisar classificação', 'Não classificado', 'Descrição nova: revisão manual necessária.'))
        if k in vistos or (k in DUPLICADOS and DUPLICADOS[k] in chaves):
            destino = DUPLICADOS.get(k, k)
            excluidos.append({**r, 'Categoria': c, 'Situação da premiação': s, 'Chave mantida': ' | '.join(destino)})
            continue
        vistos.add(k)
        detalhado.append({**r, 'Categoria': c, 'Situação da premiação': s, 'Nota de classificação': nota})
    gravar(saida/'classificacao_detalhada.csv', detalhado, campos+['Categoria','Situação da premiação','Nota de classificação'])
    gravar(saida/'duplicidades_excluidas.csv', excluidos, campos+['Categoria','Situação da premiação','Chave mantida'])
    por_ano = Counter((r['Categoria'], r['Ano']) for r in detalhado)
    por_situacao = Counter((r['Categoria'],r['Situação da premiação']) for r in detalhado)
    por_professor = Counter((r['Ano'],r['Professor / nome'],r['Categoria'],r['Situação da premiação']) for r in detalhado)
    anos = [str(y) for y in range(2019, 2026)]
    anos = sorted(set(anos) | {r['Ano'] for r in detalhado})
    situacoes = ['Premiação / colocação','Menção honrosa','Finalista / indicação','Outro reconhecimento','Não classificado']
    total = Counter(r['Categoria'] for r in detalhado)
    resumo = []
    for c,n in sorted(total.items(), key=lambda item: (-item[1], item[0])):
        resumo.append({'Categoria':c, **{y:por_ano[c,y] for y in anos}, **{s:por_situacao[c,s] for s in situacoes}, 'Total':n})
    gravar(saida/'resumo_categorias.csv', resumo, ['Categoria']+anos+situacoes+['Total'])
    campos_prof = ['Ano','Professor / nome','Categoria','Situação da premiação','Quantidade']
    gravar(saida/'resumo_professor_ano_categoria.csv', [dict(zip(campos_prof, (*k,n))) for k,n in sorted(por_professor.items())], campos_prof)
    assert len(detalhado)+len(excluidos)==len(linhas)
    assert sum(total.values())==len(detalhado)
    print(f'{len(linhas)} registros de entrada; {len(excluidos)} duplicidades excluídas; {len(detalhado)} registros classificados.')
    print("\nRESULTADO POR CATEGORIA DE PREMIAÇÃO")
    print(f"{'Categoria':62} {'Prêmios':>8} {'Menções':>8} {'Finalistas':>10} {'Outros':>7} {'Revisar':>7} {'Total':>7}")
    for r in resumo:
        print(f"{r['Categoria']:62} " + " ".join(f"{r[s]:>{w}}" for s,w in zip(situacoes, (8,8,10,7,7))) + f" {r['Total']:7}")
    contagem_situacoes = Counter(r['Situação da premiação'] for r in detalhado)
    print(f"{'TOTAL':62} " + " ".join(f"{contagem_situacoes[s]:>{w}}" for s,w in zip(situacoes, (8,8,10,7,7))) + f" {len(detalhado):7}")
    print(f"\nCategorias: {len(total)} | Docentes: {len({r['Professor / nome'] for r in detalhado})}")
    if contagem_situacoes['Não classificado']:
        print('ATENÇÃO: há registros novos a revisar em classificacao_detalhada.csv.')
    print(f"Relatórios por categoria: {saida.resolve()}")
    return detalhado, resumo

def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--entrada", type=Path, default=Path("."))
    p.add_argument("--saida", type=Path, default=Path("resultado"))
    args = p.parse_args()
    args.saida.mkdir(parents=True, exist_ok=True)
    raw = coletar(args.entrada)
    dados = consolidar(raw)
    principal, fora = separar(dados)
    gravar_csv(args.saida/"premios_docentes_ic.csv",principal)
    gravar_csv(args.saida/"outros_registros_ic.csv",fora)
    gravar_csv(args.saida/"todos_registros_ic.csv",dados)
    resumo = []
    for ano in range(2019,2026):
        a = [r for r in principal if r["Ano"] == ano]
        resumo.append({"Ano":ano,"Internos":sum(r["Tipo"]=="Interno" for r in a),
                       "Externos":sum(r["Tipo"]=="Externo" for r in a),"Total":len(a),
                       "Docentes":len({r["Professor / nome"] for r in a})})
    gravar_csv(args.saida/"resumo_anual.csv",resumo)
    por_professor = collections.Counter((r["Ano"],r["Professor / nome"]) for r in principal)
    gravar_csv(args.saida/"resumo_professor_ano.csv",
               [{"Ano":a,"Professor":n,"Total":q} for (a,n),q in sorted(por_professor.items(),key=lambda item:(item[0][0],nome_key(item[0][1])))])
    print(f"Registros IC: {len(raw)} | Após consolidação: {len(dados)} | "
          f"Docentes: {len(principal)} | Excluídos pelo critério docente: {len(fora)}")
    for r in resumo:
        print(r)
    classificar_premios(args.saida / "premios_docentes_ic.csv",
                       args.saida / "categorias_ic")

if __name__ == "__main__":
    main()
