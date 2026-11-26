#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera a planilha de classificação das defesas do IC/UNICAMP para um ano de defesas.

Fluxo:
1) lê a aba do ano no arquivo "Histórico de Defesas de Mestrado e Doutorado.xlsx";
2) segue o hyperlink da coluna Data para cada página de defesa do IC;
3) extrai orientador/coorientador, titulares, suplentes e afiliações;
4) descobre todas as edições disponíveis do Anuário de Pesquisa e cruza pessoas vinculadas à UNICAMP com prêmios/homenagens de todas elas;
5) grava a saída no formato da planilha "Classificação.xlsx";
6) deixa Gênero e Raça em branco e cria uma aba "Auditoria" com avisos/erros.

Exemplo:
    python3 classificar_defesas.py \
      --historico "Histórico de Defesas de Mestrado e Doutorado.xlsx" \
      --template "Classificação.xlsx" \
      --ano 2023 \
      --saida "Classificação_2023_preenchida.xlsx"

Dependências:
    pip install openpyxl requests beautifulsoup4
"""

from __future__ import annotations

import argparse
import copy
import difflib
import html
import logging
import re
import sys
import time
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional
from urllib.parse import quote, urljoin, urlparse

import openpyxl
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


IC_BASE = "https://ic.unicamp.br"
ANUARIO_BASE = "https://pesquisa.dados.unicamp.br"
PREMIOS_MAIN = f"{ANUARIO_BASE}/premios"
PREMIOS_CATEGORIA = "PRÊMIOS E HOMENAGENS"

MAX_TITULARES = 5
MAX_SUPLENTES = 3

OUTPUT_HEADERS = ["Classificação", "Aluno"]
for role in ["Orientador", "Co-orientador"] + [f"Titular {i}" for i in range(1, 6)] + [f"Suplente {i}" for i in range(1, 4)]:
    OUTPUT_HEADERS.extend([role, "Instituto", "Prêmio", "Gênero", "Raça"])

STOP_MARKERS = {
    "noticias", "notícias", "parceiros", "graduacao", "graduação", "pos-graduacao",
    "pós-graduação", "pesquisa", "extensao", "extensão", "compartilhe", "share",
}


@dataclass
class Member:
    name: str = ""
    affiliation: str = ""


@dataclass
class Defense:
    classification: str
    student: str
    advisor_raw_excel: str
    date_text: str
    url: str
    row_number: int


@dataclass
class DefensePage:
    advisor_raw: str = ""
    advisor: str = ""
    coadvisor: str = ""
    titulares: list[Member] | None = None
    suplentes: list[Member] | None = None

    def __post_init__(self):
        self.titulares = self.titulares or []
        self.suplentes = self.suplentes or []


@dataclass
class PrizeRecord:
    person: str
    unit: str
    year: Optional[int]
    title: str
    source_url: str
    edition_year: Optional[int] = None


class PrizeIndex:
    def __init__(self):
        self.records: list[PrizeRecord] = []
        self.by_exact: dict[str, list[PrizeRecord]] = defaultdict(list)

    def add(self, rec: PrizeRecord) -> None:
        if not rec.person.strip() or not rec.title.strip():
            return
        key = normalize_name(rec.person)
        if not key:
            return
        signature = (key, normalize_text(rec.unit), rec.year, normalize_text(rec.title))
        for old in self.by_exact.get(key, []):
            old_sig = (key, normalize_text(old.unit), old.year, normalize_text(old.title))
            if signature == old_sig:
                return
        self.records.append(rec)
        self.by_exact[key].append(rec)

    def lookup(
        self,
        person: str,
        unit: str = "",
        fuzzy_threshold: float = 0.93,
    ) -> tuple[list[PrizeRecord], str, float]:
        """Retorna (registros, modo_match, score)."""
        key = normalize_name(person)
        if not key:
            return [], "sem_nome", 0.0

        def eligible(items: Iterable[PrizeRecord]) -> list[PrizeRecord]:
            out = []
            nunit = normalize_text(unit)
            for rec in items:
                if nunit and rec.unit:
                    ru = normalize_text(rec.unit)
                    if ru != nunit and nunit not in ru and ru not in nunit:
                        continue
                out.append(rec)
            return out

        exact = eligible(self.by_exact.get(key, []))
        if exact:
            return exact, "exato", 1.0

        parts = key.split()
        last = parts[-1] if parts else ""
        best_key = ""
        best_score = 0.0
        for candidate in self.by_exact.keys():
            cparts = candidate.split()
            if last and cparts and cparts[-1] != last:
                continue
            score = difflib.SequenceMatcher(None, key, candidate).ratio()
            if score > best_score:
                best_score = score
                best_key = candidate
        if best_key and best_score >= fuzzy_threshold:
            found = eligible(self.by_exact.get(best_key, []))
            if found:
                return found, "fuzzy", best_score
        return [], "nao_encontrado", best_score


# ---------- utilidades ----------

def normalize_text(value: object) -> str:
    s = "" if value is None else str(value)
    s = html.unescape(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower().strip()
    s = re.sub(r"\s+", " ", s)
    return s


def normalize_name(value: object) -> str:
    s = normalize_text(value)
    s = re.sub(r"\([^)]*\)", " ", s)
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def clean_space(value: object) -> str:
    s = "" if value is None else str(value)
    s = html.unescape(s).replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip()


def excel_safe(value: object) -> object:
    if not isinstance(value, str):
        return value
    if value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def same_person(a: str, b: str) -> bool:
    na, nb = normalize_name(a), normalize_name(b)
    if not na or not nb:
        return False
    if na == nb:
        return True
    pa, pb = na.split(), nb.split()
    if pa and pb and pa[-1] == pb[-1]:
        return difflib.SequenceMatcher(None, na, nb).ratio() >= 0.93
    return False


def make_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=2,
        connect=2,
        read=2,
        backoff_factor=1.0,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET", "POST"}),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=10, pool_maxsize=10)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
        ),
        "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.7",
    })
    return session


def fetch_text(session: requests.Session, url: str, timeout: tuple[int, int] | int = (10, 20)) -> str:
    r = session.get(url, timeout=timeout)
    r.raise_for_status()
    r.encoding = r.apparent_encoding or r.encoding
    return r.text


# ---------- leitura das defesas ----------

def read_defenses(history_path: Path, year: int) -> list[Defense]:
    wb = openpyxl.load_workbook(history_path, data_only=False, read_only=False)
    sheet_name = str(year)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"A aba {sheet_name!r} não existe. Abas disponíveis: {wb.sheetnames}")
    ws = wb[sheet_name]

    current_classification = ""
    defenses: list[Defense] = []
    for row_idx in range(1, ws.max_row + 1):
        col_a = clean_space(ws.cell(row_idx, 1).value)
        normalized_a = normalize_text(col_a)
        if normalized_a in {"mestrado", "doutorado", "mestrado profissional"}:
            current_classification = {
                "mestrado": "Mestrado",
                "doutorado": "Doutorado",
                "mestrado profissional": "Mestrado Profissional",
            }[normalized_a]
            continue

        # Cabeçalhos e linhas vazias.
        if normalize_text(ws.cell(row_idx, 3).value) == "data":
            continue

        date_cell = ws.cell(row_idx, 3)
        student = clean_space(ws.cell(row_idx, 5).value)
        advisor = clean_space(ws.cell(row_idx, 6).value)
        if not current_classification or not student or date_cell.value is None:
            continue

        url = ""
        if date_cell.hyperlink and date_cell.hyperlink.target:
            url = clean_space(date_cell.hyperlink.target)
        elif isinstance(date_cell.value, str) and date_cell.value.startswith("http"):
            url = date_cell.value.strip()

        if hasattr(date_cell.value, "strftime"):
            date_text = date_cell.value.strftime("%Y-%m-%d")
        else:
            date_text = clean_space(date_cell.value)

        defenses.append(Defense(
            classification=current_classification,
            student=student,
            advisor_raw_excel=advisor,
            date_text=date_text,
            url=url,
            row_number=row_idx,
        ))
    return defenses


# ---------- parsing do site do IC ----------

def split_advisors(raw: str) -> tuple[str, str]:
    """Separa orientador e coorientador em formatos encontrados no histórico/site."""
    s = clean_space(raw)
    if not s:
        return "", ""

    s = re.sub(r"^orientador(?:a)?\s*:\s*", "", s, flags=re.I)
    # Formas: "X - Coorientador: Y", "X / Coorientadora - Y", etc.
    m = re.split(
        r"\s*(?:/|\-|–|—)?\s*co-?orientador(?:a)?\s*(?:\:|\-|–|—)\s*",
        s,
        maxsplit=1,
        flags=re.I,
    )
    if len(m) == 2:
        return clean_space(m[0]), clean_space(m[1])

    # Forma com "Coorientador:" sem separador anterior.
    m2 = re.split(r"\s*co-?orientador(?:a)?\s*:\s*", s, maxsplit=1, flags=re.I)
    if len(m2) == 2:
        first = re.sub(r"[\-/–—]+$", "", m2[0]).strip()
        return clean_space(first), clean_space(m2[1])

    # Cotutela: no histórico de 2023 há "A e B (cotutela)".
    if re.search(r"\bcotutela\b", s, flags=re.I):
        s2 = re.sub(r"\s*\(\s*cotutela\s*\)\s*$", "", s, flags=re.I).strip()
        parts = re.split(r"\s+e\s+", s2, maxsplit=1, flags=re.I)
        if len(parts) == 2:
            return clean_space(parts[0]), clean_space(parts[1])

    return s, ""


def looks_like_affiliation(s: str) -> bool:
    raw = clean_space(s)
    n = normalize_text(raw)
    if not raw:
        return False
    if "/" in raw:
        return True
    if "unicamp" in n or "universidade" in n or "university" in n or "instituto" in n:
        return True
    if len(raw) <= 24 and len(raw.split()) <= 3:
        if raw.isupper() or len(raw.split()) == 1:
            return True
    return False


def split_member_line(line: str) -> Optional[Member]:
    s = clean_space(line)
    if not s:
        return None
    for sep in ("|", "\t"):
        if sep in s:
            left, right = s.split(sep, 1)
            if clean_space(left):
                return Member(clean_space(left), clean_space(right))
    # Em alguns temas do WordPress a separação vira vários espaços.
    parts = re.split(r"\s{2,}", s, maxsplit=1)
    if len(parts) == 2 and looks_like_affiliation(parts[1]):
        return Member(clean_space(parts[0]), clean_space(parts[1]))
    return None


def parse_member_block(lines: list[str]) -> list[Member]:
    cleaned = [clean_space(x) for x in lines if clean_space(x)]
    out: list[Member] = []
    i = 0
    while i < len(cleaned):
        line = cleaned[i]
        explicit = split_member_line(line)
        if explicit:
            out.append(explicit)
            i += 1
            continue
        if looks_like_affiliation(line):
            i += 1
            continue
        if i + 1 < len(cleaned) and looks_like_affiliation(cleaned[i + 1]):
            out.append(Member(line, cleaned[i + 1]))
            i += 2
            continue
        # Último fallback: guarda o nome sem afiliação.
        out.append(Member(line, ""))
        i += 1
    return out


def _marker_index(lines: list[str], candidates: set[str], start: int = 0) -> int:
    for i in range(start, len(lines)):
        n = normalize_text(lines[i]).rstrip(":")
        if n in candidates:
            return i
    return -1


def parse_defense_page(page_html: str) -> DefensePage:
    soup = BeautifulSoup(page_html, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg"]):
        tag.decompose()
    lines = [clean_space(x) for x in soup.stripped_strings if clean_space(x)]

    advisor_raw = ""
    for i, line in enumerate(lines):
        if normalize_text(line).rstrip(":") in {"orientador / docente", "orientador/docente", "orientador", "orientadora / docente"}:
            if i + 1 < len(lines):
                advisor_raw = lines[i + 1]
                break
    advisor, coadvisor = split_advisors(advisor_raw)

    banca_idx = _marker_index(lines, {"banca examinadora"})
    titulares: list[Member] = []
    suplentes: list[Member] = []
    if banca_idx >= 0:
        tit_idx = _marker_index(lines, {"titulares", "titular"}, banca_idx + 1)
        sup_idx = _marker_index(lines, {"suplentes", "suplente"}, (tit_idx + 1 if tit_idx >= 0 else banca_idx + 1))
        if tit_idx >= 0:
            tit_end = sup_idx if sup_idx >= 0 else len(lines)
            titulares = parse_member_block(lines[tit_idx + 1:tit_end])
        if sup_idx >= 0:
            sup_end = len(lines)
            for j in range(sup_idx + 1, len(lines)):
                if normalize_text(lines[j]).lstrip("#").strip().rstrip(":") in STOP_MARKERS:
                    sup_end = j
                    break
            suplentes = parse_member_block(lines[sup_idx + 1:sup_end])

    return DefensePage(
        advisor_raw=advisor_raw,
        advisor=advisor,
        coadvisor=coadvisor,
        titulares=titulares,
        suplentes=suplentes,
    )


# ---------- afiliações ----------

def affiliation_for_person(name: str, members: Iterable[Member]) -> str:
    if not name:
        return ""
    for m in members:
        if same_person(name, m.name):
            return m.affiliation
    return ""


def is_unicamp_affiliation(affiliation: str) -> bool:
    return "unicamp" in normalize_text(affiliation)


def unicamp_unit_from_affiliation(affiliation: str) -> str:
    s = clean_space(affiliation)
    if not s:
        return ""
    parts = [p.strip() for p in s.split("/") if p.strip()]
    for idx, p in enumerate(parts):
        if normalize_text(p) == "unicamp" and idx > 0:
            return re.sub(r"[^A-Za-z0-9_-]", "", parts[idx - 1].upper())
    if normalize_text(s).endswith("unicamp") and parts:
        return re.sub(r"[^A-Za-z0-9_-]", "", parts[0].upper())
    return ""


# ---------- seleção das edições do Anuário ----------

def selected_anuario_year(text: str) -> Optional[int]:
    plain = clean_space(BeautifulSoup(text, "html.parser").get_text(" ", strip=True))
    m = re.search(r"\bBem[- ]?vindo\s*\(\s*((?:19|20)\d{2})\s*\)", plain, flags=re.I)
    return int(m.group(1)) if m else None


def page_mentions_selected_year(text: str, year: int) -> bool:
    return selected_anuario_year(text) == year


def discover_anuario_years(session: requests.Session, logger: logging.Logger) -> tuple[list[int], Optional[int]]:
    """Descobre as edições anuais disponíveis no seletor do Anuário.

    Prioridade:
      1) opções <option> de selects da página;
      2) links cujo texto seja somente um ano;
      3) fallback baseado no texto oficial "a partir de 2019" e no ano selecionado.
    """
    r = session.get(ANUARIO_BASE + "/", timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    current = selected_anuario_year(r.text)

    years: set[int] = set()
    for select in soup.find_all("select"):
        option_years = set()
        for opt in select.find_all("option"):
            txt = clean_space(opt.get_text())
            if re.fullmatch(r"(?:19|20)\d{2}", txt):
                option_years.add(int(txt))
        if len(option_years) >= 2:
            years.update(option_years)

    if not years:
        for a in soup.find_all("a"):
            txt = clean_space(a.get_text())
            if re.fullmatch(r"(?:19|20)\d{2}", txt):
                years.add(int(txt))

    if not years and current is not None:
        plain = normalize_text(soup.get_text(" ", strip=True))
        m = re.search(r"a partir de\s+((?:19|20)\d{2})", plain)
        start = int(m.group(1)) if m else 2019
        if start <= current:
            years.update(range(start, current + 1))

    if current is not None:
        years.add(current)

    out = sorted(y for y in years if 1900 <= y <= 2100)
    if not out:
        raise RuntimeError("Não foi possível descobrir nenhuma edição anual disponível no Anuário.")
    logger.info("Edições do Anuário descobertas: %s", ", ".join(map(str, out)))
    return out, current


def parse_year_spec(spec: str, available_years: list[int]) -> list[int]:
    """Interpreta 'auto', '2019-2024' ou '2019,2021,2024'."""
    spec = clean_space(spec).lower()
    if spec in {"", "auto", "todos", "all"}:
        return list(available_years)
    requested: set[int] = set()
    for part in re.split(r"[,;]", spec):
        part = part.strip()
        if not part:
            continue
        m = re.fullmatch(r"((?:19|20)\d{2})\s*-\s*((?:19|20)\d{2})", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            lo, hi = sorted((a, b))
            requested.update(range(lo, hi + 1))
            continue
        if re.fullmatch(r"(?:19|20)\d{2}", part):
            requested.add(int(part))
            continue
        raise ValueError(f"Especificação de anos inválida: {part!r}")
    return sorted(y for y in requested if y in set(available_years))


def try_select_anuario_year(session: requests.Session, year: int, logger: logging.Logger) -> bool:

    try:
        r = session.get(ANUARIO_BASE + "/", timeout=30)
        r.raise_for_status()
    except Exception as exc:
        logger.warning("Não foi possível abrir o Anuário para selecionar o ano: %s", exc)
        return False

    if page_mentions_selected_year(r.text, year):
        return True

    soup = BeautifulSoup(r.text, "html.parser")
    year_str = str(year)

    for select in soup.find_all("select"):
        option = next((o for o in select.find_all("option") if clean_space(o.get_text()) == year_str), None)
        if option is None:
            continue
        value = option.get("value", year_str)
        name = select.get("name") or select.get("id")
        form = select.find_parent("form")
        if form and name:
            data = {}
            for inp in form.find_all("input"):
                inp_name = inp.get("name")
                if inp_name:
                    data[inp_name] = inp.get("value", "")
            data[name] = value
            action = urljoin(r.url, form.get("action") or r.url)
            method = (form.get("method") or "get").lower()
            try:
                rr = session.post(action, data=data, timeout=30) if method == "post" else session.get(action, params=data, timeout=30)
                if rr.ok and page_mentions_selected_year(rr.text, year):
                    logger.info("Edição %s selecionada no Anuário via formulário.", year)
                    return True
                # Alguns formulários redirecionam de volta à home e atualizam cookie/sessão.
                check = session.get(ANUARIO_BASE + "/", timeout=30)
                if check.ok and page_mentions_selected_year(check.text, year):
                    logger.info("Edição %s selecionada no Anuário via sessão.", year)
                    return True
            except requests.RequestException:
                pass

        if name:
            for method in ("get", "post"):
                try:
                    if method == "get":
                        rr = session.get(r.url, params={name: value}, timeout=30)
                    else:
                        rr = session.post(r.url, data={name: value}, timeout=30)
                    if rr.ok:
                        check = session.get(ANUARIO_BASE + "/", timeout=30)
                        if page_mentions_selected_year(rr.text, year) or (check.ok and page_mentions_selected_year(check.text, year)):
                            logger.info("Edição %s selecionada usando o campo %s do seletor.", year, name)
                            return True
                except requests.RequestException:
                    pass

        onchange = clean_space(select.get("onchange"))
        if onchange:
            for quoted in re.findall(r"['\"]([^'\"]+)['\"]", onchange):
                candidate = quoted.replace("{year}", year_str).replace("YEAR", year_str)
                if year_str in candidate or str(value) in candidate:
                    try:
                        rr = session.get(urljoin(r.url, candidate), timeout=30)
                        if rr.ok:
                            check = session.get(ANUARIO_BASE + "/", timeout=30)
                            if page_mentions_selected_year(rr.text, year) or (check.ok and page_mentions_selected_year(check.text, year)):
                                logger.info("Edição %s selecionada via onchange do seletor.", year)
                                return True
                    except requests.RequestException:
                        pass

        if value and (str(value).startswith("http") or str(value).startswith("/")):
            try:
                rr = session.get(urljoin(r.url, str(value)), timeout=30)
                if rr.ok:
                    check = session.get(ANUARIO_BASE + "/", timeout=30)
                    if page_mentions_selected_year(check.text, year) or page_mentions_selected_year(rr.text, year):
                        logger.info("Edição %s selecionada no Anuário via URL do seletor.", year)
                        return True
            except requests.RequestException:
                pass

    for a in soup.find_all("a"):
        if clean_space(a.get_text()) != year_str:
            continue
        href = a.get("href")
        if not href:
            continue
        try:
            rr = session.get(urljoin(r.url, href), timeout=30)
            if rr.ok:
                check = session.get(ANUARIO_BASE + "/", timeout=30)
                if page_mentions_selected_year(rr.text, year) or page_mentions_selected_year(check.text, year):
                    logger.info("Edição %s selecionada no Anuário via link do menu.", year)
                    return True
        except requests.RequestException:
            pass

    for key in ("ano", "year", "edicao", "edition"):
        for method in ("get", "post"):
            try:
                if method == "get":
                    rr = session.get(ANUARIO_BASE + "/", params={key: year_str}, timeout=30)
                else:
                    rr = session.post(ANUARIO_BASE + "/", data={key: year_str}, timeout=30)
                if rr.ok:
                    check = session.get(ANUARIO_BASE + "/", timeout=30)
                    if page_mentions_selected_year(rr.text, year) or page_mentions_selected_year(check.text, year):
                        logger.info("Edição %s selecionada no Anuário com parâmetro %s.", year, key)
                        return True
            except requests.RequestException:
                pass

    logger.warning(
        "Não consegui confirmar automaticamente a edição %s do Anuário. "
        "O script continuará, mas não marcará ausência de prêmio como 'Não'; usará 'Não verificado'.",
        year,
    )
    return False


# ---------- parsing de prêmios ----------

def table_headers(table) -> list[str]:
    first = table.find("tr")
    if not first:
        return []
    return [clean_space(c.get_text(" ", strip=True)) for c in first.find_all(["th", "td"])]


def find_col(headers: list[str], candidates: Iterable[str]) -> int:
    nh = [normalize_text(h) for h in headers]
    for cand in candidates:
        c = normalize_text(cand)
        for i, h in enumerate(nh):
            if h == c or c in h:
                return i
    return -1


def extract_author_names(cell) -> list[str]:
    chunks = [clean_space(x) for x in cell.stripped_strings if clean_space(x)]
    names: list[str] = []
    for chunk in chunks:
        pieces = re.split(r"\s*[;\n]+\s*", chunk)
        for p in pieces:
            p = clean_space(p)
            if not p:
                continue
            p = re.sub(r"\s*\([^)]*(?:docente|pesquisador|estudante|pos|pós|unicamp|/)[^)]*\)\s*", "", p, flags=re.I)
            p = clean_space(p)
            if p and len(p.split()) >= 2 and normalize_text(p) not in {"autores da unicamp", "autores"}:
                names.append(p)
    seen = set()
    out = []
    for n in names:
        k = normalize_name(n)
        if k and k not in seen:
            seen.add(k)
            out.append(n)
    return out


def parse_year(value: str) -> Optional[int]:
    m = re.search(r"\b(19|20)\d{2}\b", clean_space(value))
    return int(m.group(0)) if m else None


def scrape_prizes_from_html(page_html: str, source_url: str, forced_unit: str = "", edition_year: Optional[int] = None) -> list[PrizeRecord]:
    soup = BeautifulSoup(page_html, "html.parser")
    records: list[PrizeRecord] = []
    for table in soup.find_all("table"):
        headers = table_headers(table)
        if not headers:
            continue
        person_col = find_col(headers, ["Autores da Unicamp", "Nome", "Pessoa", "Docente", "Pesquisador", "Autor"])
        title_col = find_col(headers, ["Título", "Premio", "Prêmio", "Homenagem", "Descrição"])
        year_col = find_col(headers, ["Ano"])
        unit_col = find_col(headers, ["Unidade", "Instituto", "Sigla"])
        if person_col < 0 or title_col < 0:
            continue

        rows = table.find_all("tr")[1:]
        for tr in rows:
            cells = tr.find_all(["td", "th"])
            if max(person_col, title_col) >= len(cells):
                continue
            title = clean_space(cells[title_col].get_text(" ", strip=True))
            if not title:
                continue
            year = parse_year(cells[year_col].get_text(" ", strip=True)) if year_col >= 0 and year_col < len(cells) else None
            unit = forced_unit
            if unit_col >= 0 and unit_col < len(cells):
                unit = clean_space(cells[unit_col].get_text(" ", strip=True)) or forced_unit
            names = extract_author_names(cells[person_col])
            if not names:
                raw_person = clean_space(cells[person_col].get_text(" ", strip=True))
                raw_person = re.sub(r"\s*\([^)]*\)\s*", "", raw_person).strip()
                if raw_person and len(raw_person.split()) >= 2:
                    names = [raw_person]
            for person in names:
                records.append(PrizeRecord(person=person, unit=unit, year=year, title=title, source_url=source_url, edition_year=edition_year))
    return records


def unit_prize_url(unit: str) -> str:
    unit = re.sub(r"[^A-Za-z0-9_-]", "", unit.upper())
    category = quote(PREMIOS_CATEGORIA, safe="")
    return f"{ANUARIO_BASE}/lattes/unidades/premios_e_homenagens/{unit}/{category}"


def load_unit_prizes_for_year(
    session: requests.Session,
    prize_index: PrizeIndex,
    unit: str,
    edition_year: int,
    loaded_unit_years: set[tuple[str, int]],
    logger: logging.Logger,
    cache_dir: Optional[Path] = None,
    refresh_cache: bool = False,
) -> bool:

    unit = unit.upper().strip()
    key = (unit, edition_year)
    if not unit:
        return False
    if key in loaded_unit_years:
        return True
    loaded_unit_years.add(key)
    url = unit_prize_url(unit)

    cache_path: Optional[Path] = None
    if cache_dir is not None:
        cache_dir.mkdir(parents=True, exist_ok=True)
        cache_path = cache_dir / f"premios_{edition_year}_{unit}.html"

    page_html = ""
    from_cache = False
    if cache_path is not None and cache_path.exists() and not refresh_cache:
        try:
            page_html = cache_path.read_text(encoding="utf-8")
            from_cache = True
            logger.info("Prêmios: edição %s / unidade %s -> usando cache local.", edition_year, unit)
        except OSError as exc:
            logger.warning("Não foi possível ler o cache %s: %s", cache_path, exc)
            page_html = ""

    if not page_html:
        try:
            page_html = fetch_text(session, url)
        except Exception as exc:
            logger.warning(
                "Não foi possível ler prêmios da unidade %s na edição %s (%s): %s",
                unit, edition_year, url, exc,
            )
            return False

    selected = selected_anuario_year(page_html)
    if selected is not None and selected != edition_year:
        logger.warning(
            "Unidade %s: página respondeu na edição %s, mas era esperada %s; ignorando esta resposta.",
            unit, selected, edition_year,
        )
        if from_cache and cache_path is not None:
            try:
                cache_path.unlink(missing_ok=True)
            except OSError:
                pass
        return False

    recs = scrape_prizes_from_html(page_html, url, forced_unit=unit, edition_year=edition_year)
    for rec in recs:
        prize_index.add(rec)

    if cache_path is not None and not from_cache:
        try:
            cache_path.write_text(page_html, encoding="utf-8")
        except OSError as exc:
            logger.warning("Não foi possível gravar o cache %s: %s", cache_path, exc)

    logger.info("Prêmios: edição %s / unidade %s -> %d registros lidos.", edition_year, unit, len(recs))
    return True


class PrizeCatalog:

    def __init__(
        self,
        years: list[int],
        logger: logging.Logger,
        cache_dir: Optional[Path] = None,
        refresh_cache: bool = False,
    ):
        self.years = list(years)
        self.logger = logger
        self.index = PrizeIndex()
        self.sessions: dict[int, requests.Session] = {}
        self.confirmed_years: list[int] = []
        self.unconfirmed_years: list[int] = []
        self.loaded_unit_years: set[tuple[str, int]] = set()
        self.successful_unit_years: set[tuple[str, int]] = set()
        self.failed_unit_years: set[tuple[str, int]] = set()
        self.cache_dir = cache_dir
        self.refresh_cache = refresh_cache

        for year in self.years:
            sess = make_session()
            if try_select_anuario_year(sess, year, logger):
                self.sessions[year] = sess
                self.confirmed_years.append(year)
            else:
                self.unconfirmed_years.append(year)

        if not self.confirmed_years:
            raise RuntimeError(
                "Nenhuma edição do Anuário pôde ser selecionada/confirmada. "
                "Sem isso, o cruzamento de prêmios não seria confiável."
            )

        logger.info("Edições confirmadas para consulta: %s", ", ".join(map(str, self.confirmed_years)))
        if self.unconfirmed_years:
            logger.warning("Edições que não puderam ser confirmadas: %s", ", ".join(map(str, self.unconfirmed_years)))

    def load_unit(self, unit: str) -> None:
        unit = unit.upper().strip()
        if not unit:
            return
        for year in self.confirmed_years:
            key = (unit, year)
            if key in self.successful_unit_years or key in self.failed_unit_years:
                continue
            ok = load_unit_prizes_for_year(
                self.sessions[year],
                self.index,
                unit,
                year,
                self.loaded_unit_years,
                self.logger,
                cache_dir=self.cache_dir,
                refresh_cache=self.refresh_cache,
            )
            if ok:
                self.successful_unit_years.add(key)
            else:
                self.failed_unit_years.add(key)

    def lookup(self, person: str, unit: str = "") -> tuple[list[PrizeRecord], str, float]:
        if unit:
            self.load_unit(unit)
        return self.index.lookup(person, unit=unit)

    def verification_years_for_unit(self, unit: str) -> tuple[list[int], list[int]]:
        unit = unit.upper().strip()
        ok = sorted(y for (u, y) in self.successful_unit_years if u == unit)
        failed = sorted(y for (u, y) in self.failed_unit_years if u == unit)
        failed = sorted(set(failed) | set(self.unconfirmed_years))
        return ok, failed


def format_year_list(years: Iterable[int]) -> str:
    vals = sorted(set(int(y) for y in years))
    if not vals:
        return ""
    if vals == list(range(vals[0], vals[-1] + 1)) and len(vals) >= 3:
        return f"{vals[0]}–{vals[-1]}"
    return ", ".join(map(str, vals))


# ---------- montagem da saída ----------

def prize_text_for_person(
    person: str,
    affiliation: str,
    catalog: PrizeCatalog,
    logger: logging.Logger,
) -> tuple[str, str, float]:
    if not person:
        return "", "sem_nome", 0.0
    if affiliation and not is_unicamp_affiliation(affiliation):
        return "N/A (fora da Unicamp)", "externo", 1.0

    unit = unicamp_unit_from_affiliation(affiliation)
    if not unit:
        return "Não verificado — instituto/unidade não identificado", "sem_unidade", 0.0

    recs, mode, score = catalog.lookup(person, unit=unit)
    if recs:
        unique: dict[tuple[int, str], PrizeRecord] = {}
        for rec in recs:
            y = rec.year or rec.edition_year or 0
            unique[(y, normalize_text(rec.title))] = rec
        ordered = sorted(unique.values(), key=lambda r: ((r.year or r.edition_year or 0), normalize_text(r.title)))
        labels = []
        for rec in ordered:
            y = rec.year or rec.edition_year
            labels.append(f"{y} — {rec.title}" if y else rec.title)
        return " | ".join(labels), mode, score

    verified_years, failed_years = catalog.verification_years_for_unit(unit)
    verified = format_year_list(verified_years)
    missing = format_year_list(failed_years)
    if failed_years:
        if verified_years:
            return (
                f"Não localizado nos anos verificados ({verified}); "
                f"não foi possível verificar: {missing}",
                mode, score,
            )
        return f"Não verificado — falha ao consultar a unidade {unit} nos anos: {missing}", mode, score
    if verified_years:
        return f"Não localizado nas edições consultadas ({verified})", mode, score
    return f"Não verificado — nenhuma edição da unidade {unit} foi carregada", mode, score


def person_block_values(
    member: Member,
    catalog: PrizeCatalog,
    logger: logging.Logger,
) -> tuple[list[str], str, float]:
    prize, mode, score = prize_text_for_person(member.name, member.affiliation, catalog, logger)
    return [member.name, member.affiliation, prize, "", ""], mode, score


def prepare_output_workbook(template_path: Optional[Path]) -> openpyxl.Workbook:
    if template_path and template_path.exists():
        wb = openpyxl.load_workbook(template_path)
        ws = wb[wb.sheetnames[0]]
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        for col, header in enumerate(OUTPUT_HEADERS, start=1):
            ws.cell(1, col).value = header
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Classificação"
        for col, header in enumerate(OUTPUT_HEADERS, start=1):
            ws.cell(1, col).value = header
    return wb


def style_output(ws) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(OUTPUT_HEADERS))}{max(ws.max_row, 1)}"
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in ws[1][:len(OUTPUT_HEADERS)]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
    widths = {1: 16, 2: 34}
    for col in range(3, len(OUTPUT_HEADERS) + 1):
        header = clean_space(ws.cell(1, col).value)
        widths[col] = 32 if header in {"Orientador", "Co-orientador"} or header.startswith(("Titular", "Suplente")) else (24 if header == "Instituto" else 46 if header == "Prêmio" else 14)
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


def process(
    history_path: Path,
    template_path: Optional[Path],
    output_path: Path,
    year: int,
    prizes_year_spec: str,
    delay: float,
    logger: logging.Logger,
    cache_dir: Optional[Path] = None,
    refresh_cache: bool = False,
) -> None:
    defenses = read_defenses(history_path, year)
    if not defenses:
        raise RuntimeError(f"Nenhuma defesa foi encontrada na aba {year}.")
    logger.info("%d defesas encontradas na aba %s.", len(defenses), year)

    session_ic = make_session()

    probe_session = make_session()
    available_prize_years, _current_prize_year = discover_anuario_years(probe_session, logger)
    selected_prize_years = parse_year_spec(prizes_year_spec, available_prize_years)
    if not selected_prize_years:
        raise RuntimeError(
            f"Nenhuma edição corresponde a --premios-anos={prizes_year_spec!r}. "
            f"Disponíveis: {', '.join(map(str, available_prize_years))}"
        )
    catalog = PrizeCatalog(
        selected_prize_years,
        logger,
        cache_dir=cache_dir,
        refresh_cache=refresh_cache,
    )

    wb = prepare_output_workbook(template_path)
    ws = wb[wb.sheetnames[0]]
    ws.title = f"Classificação {year}"

    if "Auditoria" in wb.sheetnames:
        del wb["Auditoria"]
    audit = wb.create_sheet("Auditoria")
    audit_headers = [
        "Linha histórico", "Classificação", "Aluno", "Data", "URL", "Status página",
        "Orientação (Excel)", "Orientação (site)", "Avisos", "Matches de prêmio",
        "Edições de prêmio disponíveis", "Edições confirmadas", "Edições não confirmadas"
    ]
    audit.append(audit_headers)

    for pos, defense in enumerate(defenses, start=1):
        logger.info("[%d/%d] %s — %s", pos, len(defenses), defense.classification, defense.student)
        warnings: list[str] = []
        page_status = "OK"
        parsed = DefensePage()

        if defense.url:
            try:
                page_html = fetch_text(session_ic, defense.url)
                parsed = parse_defense_page(page_html)
            except Exception as exc:
                page_status = f"ERRO: {type(exc).__name__}: {exc}"
                warnings.append("Página do IC não pôde ser processada; usados dados do XLSX quando possível.")
        else:
            page_status = "SEM URL"
            warnings.append("A célula Data não possui hyperlink.")

        if not parsed.advisor_raw:
            parsed.advisor_raw = defense.advisor_raw_excel
            parsed.advisor, parsed.coadvisor = split_advisors(defense.advisor_raw_excel)
        elif defense.advisor_raw_excel and not same_person(split_advisors(defense.advisor_raw_excel)[0], parsed.advisor):
            warnings.append("Orientador no site difere do texto do XLSX; prevaleceu o site.")

        all_members = parsed.titulares + parsed.suplentes
        advisor_aff = affiliation_for_person(parsed.advisor, all_members)
        coadvisor_aff = affiliation_for_person(parsed.coadvisor, all_members)

        if parsed.advisor and not advisor_aff:
            advisor_aff = "IC/UNICAMP"
            warnings.append("Afiliação do orientador não apareceu na banca; aplicado fallback IC/UNICAMP.")
        if parsed.coadvisor and not coadvisor_aff:
            warnings.append("Afiliação do coorientador não foi localizada na lista da banca.")

        if len(parsed.titulares) > MAX_TITULARES:
            warnings.append(f"Há {len(parsed.titulares)} titulares; a tabela guarda os primeiros {MAX_TITULARES}.")
        if len(parsed.suplentes) > MAX_SUPLENTES:
            extras = "; ".join(f"{m.name} ({m.affiliation})" for m in parsed.suplentes[MAX_SUPLENTES:])
            warnings.append(f"Há {len(parsed.suplentes)} suplentes; extras além de Suplente {MAX_SUPLENTES}: {extras}")

        advisor_member = Member(parsed.advisor, advisor_aff)
        coadvisor_member = Member(parsed.coadvisor, coadvisor_aff)
        titulares = (parsed.titulares + [Member()] * MAX_TITULARES)[:MAX_TITULARES]
        suplentes = (parsed.suplentes + [Member()] * MAX_SUPLENTES)[:MAX_SUPLENTES]

        row_values: list[object] = [defense.classification, defense.student]
        match_notes: list[str] = []
        for label, member in [
            ("Orientador", advisor_member),
            ("Co-orientador", coadvisor_member),
            *[(f"Titular {i+1}", m) for i, m in enumerate(titulares)],
            *[(f"Suplente {i+1}", m) for i, m in enumerate(suplentes)],
        ]:
            vals, mode, score = person_block_values(member, catalog, logger)
            row_values.extend(vals)
            if member.name:
                match_notes.append(f"{label}: {mode} ({score:.3f})")

        if len(row_values) != len(OUTPUT_HEADERS):
            raise AssertionError(f"Erro interno: linha possui {len(row_values)} valores; esperado {len(OUTPUT_HEADERS)}.")
        ws.append([excel_safe(v) for v in row_values])

        audit.append([
            defense.row_number,
            defense.classification,
            defense.student,
            defense.date_text,
            defense.url,
            page_status,
            defense.advisor_raw_excel,
            parsed.advisor_raw,
            " | ".join(warnings),
            " | ".join(match_notes),
            format_year_list(available_prize_years),
            format_year_list(catalog.confirmed_years),
            format_year_list(catalog.unconfirmed_years),
        ])
        if delay > 0:
            time.sleep(delay)

    style_output(ws)
    audit.freeze_panes = "A2"
    audit.auto_filter.ref = f"A1:M{audit.max_row}"
    for cell in audit[1]:
        cell.font = Font(bold=True)
    audit.column_dimensions["A"].width = 16
    audit.column_dimensions["B"].width = 18
    audit.column_dimensions["C"].width = 34
    audit.column_dimensions["D"].width = 14
    audit.column_dimensions["E"].width = 70
    audit.column_dimensions["F"].width = 36
    audit.column_dimensions["G"].width = 55
    audit.column_dimensions["H"].width = 55
    audit.column_dimensions["I"].width = 90
    audit.column_dimensions["J"].width = 70
    audit.column_dimensions["K"].width = 28
    audit.column_dimensions["L"].width = 28
    audit.column_dimensions["M"].width = 28

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    logger.info("Arquivo salvo em: %s", output_path.resolve())


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Monta a classificação das defesas do IC/UNICAMP e cruza com prêmios da Unicamp.")
    p.add_argument("--historico", required=True, type=Path, help="XLSX com o histórico de defesas.")
    p.add_argument("--template", type=Path, default=None, help="XLSX Classificação.xlsx (opcional, preserva o formato do cabeçalho).")
    p.add_argument("--saida", type=Path, default=Path("Classificação_2023_preenchida.xlsx"), help="Arquivo XLSX de saída.")
    p.add_argument("--ano", type=int, default=2023, help="Ano das defesas (default: 2023).")
    p.add_argument(
        "--premios-anos",
        default="auto",
        help=(
            "Edições do Anuário a consultar. Default: auto (todas as edições disponíveis no seletor). "
            "Também aceita faixa, ex.: 2019-2025, ou lista, ex.: 2019,2021,2024."
        ),
    )
    p.add_argument("--premios-ano", type=int, default=None, help=argparse.SUPPRESS)
    p.add_argument("--premios-todos", action="store_true", help=argparse.SUPPRESS)
    p.add_argument("--delay", type=float, default=0.25, help="Pausa entre defesas, em segundos (default: 0.25).")
    p.add_argument(
        "--cache-dir",
        type=Path,
        default=Path(".cache_premios_unicamp"),
        help="Pasta de cache das páginas de prêmios (default: .cache_premios_unicamp).",
    )
    p.add_argument(
        "--refresh-cache",
        action="store_true",
        help="Ignora o cache de prêmios e baixa novamente as páginas do Anuário.",
    )
    p.add_argument("--verbose", action="store_true", help="Mostra logs detalhados.")
    return p.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%H:%M:%S",
    )
    logger = logging.getLogger("classificar_defesas")

    if args.premios_ano is not None:
        logger.warning(
            "--premios-ano é um argumento da versão antiga e será ignorado. "
            "A versão atual consulta todas as edições disponíveis por padrão; "
            "use --premios-anos somente se quiser restringir explicitamente a faixa."
        )
    if args.premios_todos:
        logger.warning("--premios-todos é obsoleto; todas as edições já são consultadas por padrão.")

    if not args.historico.exists():
        logger.error("Arquivo de histórico não encontrado: %s", args.historico)
        return 2
    if args.template is not None and not args.template.exists():
        logger.error("Template não encontrado: %s", args.template)
        return 2

    try:
        process(
            history_path=args.historico,
            template_path=args.template,
            output_path=args.saida,
            year=args.ano,
            prizes_year_spec=args.premios_anos,
            delay=max(0.0, args.delay),
            logger=logger,
            cache_dir=args.cache_dir,
            refresh_cache=args.refresh_cache,
        )
    except KeyboardInterrupt:
        logger.error("Execução interrompida pelo usuário.")
        return 130
    except Exception as exc:
        logger.exception("Falha fatal: %s", exc)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
