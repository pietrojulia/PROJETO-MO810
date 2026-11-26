#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Atualiza, em modo protegido, os prêmios de uma planilha de defesas revisada.
Para novas defesas, --gerar-do-zero autoriza extração experimental do site.

Fluxo:
1) lê a aba do ano no arquivo "Histórico de Defesas de Mestrado e Doutorado.xlsx";
2) segue o hyperlink da coluna Data para cada página de defesa do IC;
3) extrai orientador/coorientador, titulares, suplentes e afiliações;
4) cruza membros do IC/UNICAMP com o CSV local validado de prêmios e homenagens;
5) grava a saída no formato da planilha "Classificação.xlsx";
6) deixa Gênero e Raça em branco e cria uma aba "Auditoria" com avisos/erros.

Exemplo:
    python3 classificar_defesas_premios_validados.py \
        --atualizar-classificacao "Classificação_2023_preenchida.xlsx" \
        --premios-base "/premios/resultado/categorias_ic/classificacao_detalhada.csv" \
        --ano 2023 \
        --saida "Classificação_2023_premios_validados.xlsx"

Dependências:
    pip install openpyxl requests beautifulsoup4
"""

from __future__ import annotations

import argparse
import csv
import hashlib
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

import openpyxl
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


IC_BASE = "https://ic.unicamp.br"
DEFAULT_PRIZE_BASE = Path(__file__).resolve().parent / "premios_ic_validados.csv"
CONQUISTAS = {"Premiação / colocação", "Menção honrosa"}

MAX_TITULARES = 5
MAX_SUPLENTES = 3

# Estrutura pedida, agora com Suplente 3 para todas as defesas (52 colunas úteis).
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


@dataclass(frozen=True)
class PrizeRecord:
    person: str
    year: int
    title: str
    category: str
    status: str


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
    # Timeout separado para conexão/leitura das páginas de defesa do IC.
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
    # Siglas e organizações curtas: EEP, SiDi, FSNT, Samsung etc.
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
            # Afiliação solta sem nome anterior: ignora para não deslocar toda a banca.
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
    # Remove scripts/estilos para não contaminar a sequência de texto.
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
    # Páginas que apresentam a coorientação em um campo separado.
    if not coadvisor:
        for i, line in enumerate(lines):
            if normalize_text(line).rstrip(":") in {"coorientador", "coorientadora", "co-orientador", "co-orientadora", "coorientador / docente", "co-orientador / docente"} and i+1<len(lines):
                coadvisor = lines[i+1]
                break
            m = re.match(r"^co-?orientador(?:a)?\s*:\s*(.+)$", line, re.I)
            if m:
                coadvisor = clean_space(m.group(1)); break

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


# ---------- catálogo local de prêmios validados ----------

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
    return sorted(requested)


class PrizeCatalog:
    """Consulta exclusivamente o CSV final validado. Não usa rede nem cache HTML."""

    def __init__(self, path: Path, year_spec: str, defense_year: int,
                 until_defense: bool, only_achievements: bool, logger: logging.Logger):
        self.path = path.resolve()
        if not self.path.is_file():
            raise FileNotFoundError(f"Base validada não encontrada: {self.path}. Use --premios-base.")
        self.sha256 = hashlib.sha256(self.path.read_bytes()).hexdigest()
        self.index: dict[str, list[PrizeRecord]] = defaultdict(list)
        self.all_names: dict[str, str] = {}
        records = []
        required = {"Ano", "Professor / nome", "Prêmio", "Categoria", "Situação da premiação"}
        with self.path.open(encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f, delimiter=";")
            if not required.issubset(reader.fieldnames or []):
                raise ValueError(f"CSV inválido. Esperado classificacao_detalhada.csv; faltam: {sorted(required - set(reader.fieldnames or []))}")
            seen = set()
            for row in reader:
                if not any(row.values()):
                    continue
                try:
                    rec = PrizeRecord(clean_space(row["Professor / nome"]), int(row["Ano"]),
                                      clean_space(row["Prêmio"]), clean_space(row["Categoria"]),
                                      clean_space(row["Situação da premiação"]))
                except (ValueError, TypeError) as exc:
                    raise ValueError(f"Linha {reader.line_num} inválida na base de prêmios") from exc
                key = normalize_name(rec.person)
                if not key or not rec.title or not rec.category or rec.status not in CONQUISTAS | {"Finalista / indicação", "Outro reconhecimento"}:
                    raise ValueError(f"Linha {reader.line_num} incompleta ou situação desconhecida na base de prêmios")
                sig = (key, rec.year, normalize_text(rec.title))
                if sig in seen:
                    raise ValueError(f"Duplicidade na base validada, linha {reader.line_num}: {rec.person}")
                seen.add(sig)
                records.append(rec)
                self.all_names[key] = rec.person
        if not records:
            raise ValueError("A base validada está vazia; não é possível verificar prêmios.")
        self.available_years = sorted({r.year for r in records})
        requested = parse_year_spec(year_spec, self.available_years)
        if until_defense:
            requested = [y for y in requested if y <= defense_year]
        self.confirmed_years = sorted(set(requested) & set(self.available_years))
        self.unconfirmed_years = sorted(set(requested) - set(self.available_years))
        if not self.confirmed_years:
            raise ValueError("Nenhum ano solicitado está presente na base validada.")
        self.only_achievements = only_achievements
        self.until_defense = until_defense
        self.defense_year = defense_year
        for rec in records:
            if rec.year in self.confirmed_years and (not only_achievements or rec.status in CONQUISTAS):
                self.index[normalize_name(rec.person)].append(rec)
        self.record_count = sum(map(len, self.index.values()))
        self.criteria = ("Premiações/colocações e menções honrosas" if only_achievements else
                         "Todos os registros validados (inclui finalistas e outros reconhecimentos)")
        self.criteria += (f"; até o ano da defesa ({defense_year})" if until_defense else
                          "; histórico completo dos anos selecionados (pode incluir prêmios posteriores à defesa)")
        self.coverage = "Histórico validado dos docentes do IC; correspondência por nome, independente do vínculo na banca; ausência no CSV não significa ausência de prêmio"
        logger.info("Base local: %s | %d registros originais, %d pessoas | %d registros selecionados | anos %s",
                    self.path, len(records), len(self.all_names), self.record_count, format_year_list(self.confirmed_years))
        if self.unconfirmed_years:
            logger.warning("Anos solicitados ausentes na base: %s", format_year_list(self.unconfirmed_years))
        if max(self.confirmed_years) > defense_year:
            logger.warning("A seleção inclui prêmios posteriores ao ano da defesa; use --premios-ate-defesa para limitar.")

    def lookup(self, person: str, unit: str = "IC") -> tuple[list[PrizeRecord], str, float]:
        key = normalize_name(person)
        if key in self.all_names:
            return self.index.get(key, []), "exato" if key in self.index else "exato_sem_registro_no_recorte", 1.0
        # Similaridade só é usada para sugerir revisão, jamais atribui um prêmio.
        candidates = [(difflib.SequenceMatcher(None, key, k).ratio(), name)
                      for k, name in self.all_names.items() if key and k.split()[-1] == key.split()[-1]]
        if candidates:
            score, name = max(candidates)
            if score >= 0.93:
                return [], f"revisar_nome: possível correspondência {name}", score
        return [], "nao_localizado_base_validada", 0.0

    def audit_values(self) -> list[str]:
        return [format_year_list(self.available_years), format_year_list(self.confirmed_years),
                format_year_list(self.unconfirmed_years), f"{self.path.name} | SHA256: {self.sha256}",
                self.criteria, self.coverage]


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

    # A base já foi validada quanto ao vínculo docente no ano do prêmio.
    # A afiliação na banca é uma informação diferente e não deve impedir
    # a recuperação desse histórico. Nenhuma afiliação é alterada.
    # Dois casos compostos conferidos nas planilhas revisadas: preserva o
    # texto da célula e recupera apenas os registros dos nomes confirmados.
    compostos = {
        normalize_name("Alexandre Xavier Falcão e Jean Cousty"):
            ["Alexandre Xavier Falcão", "Jean Cousty"],
        normalize_name("Leonardo Montecchi e Julio Cesar dos Reis"):
            ["Leonardo Montecchi", "Julio Cesar dos Reis"],
    }
    key = normalize_name(person)
    nomes = compostos.get(key, [person])
    matched = [normalize_name(n) for n in nomes if normalize_name(n) in catalog.all_names]
    if matched:
        recs = [r for k in dict.fromkeys(matched) for r in catalog.index.get(k, [])]
        mode = ("exato_composto: " + ", ".join(catalog.all_names[k] for k in matched)
                if key in compostos else "exato_por_nome")
        if not recs:
            return f"Não localizado na base validada do IC ({format_year_list(catalog.confirmed_years)})", mode, 1.0
        ordered = sorted(recs, key=lambda r: (r.year, normalize_text(r.title), normalize_name(r.person)))
        return " | ".join(f"{r.year} — {r.title}" for r in ordered), mode, 1.0

    # Sem correspondência exata: não atribuir prêmios por similaridade.
    _, mode, score = catalog.lookup(person)
    years = format_year_list(catalog.confirmed_years)
    if mode.startswith("revisar_nome"):
        return f"Não verificado — nome a revisar na base validada do IC ({years})", mode, score
    if affiliation and not is_unicamp_affiliation(affiliation):
        return "N/A (fora da Unicamp)", "externo_sem_registro", 1.0
    unit = unicamp_unit_from_affiliation(affiliation)
    if unit and unit != "IC":
        return f"Não verificado — unidade {unit} fora da base validada do IC", "unidade_fora_da_base", 0.0
    if not unit:
        return f"Não localizado na base validada do IC ({years}) — vínculo não identificado", "sem_unidade_sem_registro", 0.0
    return f"Não localizado na base validada do IC ({years})", mode, score


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
    prizes_base: Path = DEFAULT_PRIZE_BASE,
    until_defense: bool = False,
    only_achievements: bool = False,
) -> None:
    defenses = read_defenses(history_path, year)
    if not defenses:
        raise RuntimeError(f"Nenhuma defesa foi encontrada na aba {year}.")
    logger.info("%d defesas encontradas na aba %s.", len(defenses), year)

    session_ic = make_session()

    catalog = PrizeCatalog(prizes_base, prizes_year_spec, year, until_defense, only_achievements, logger)

    wb = prepare_output_workbook(template_path)
    ws = wb[wb.sheetnames[0]]
    ws.title = f"Classificação {year}"

    if "Auditoria" in wb.sheetnames:
        del wb["Auditoria"]
    audit = wb.create_sheet("Auditoria")
    audit_headers = [
        "Linha histórico", "Classificação", "Aluno", "Data", "URL", "Status página",
        "Orientação (Excel)", "Orientação (site)", "Avisos", "Matches de prêmio",
        "Edições de prêmio disponíveis", "Edições confirmadas", "Edições não confirmadas",
        "Fonte de prêmios", "Critério de prêmios", "Cobertura de prêmios"
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

        # Se o site omitir a coorientação, usa o histórico quando explícito e
        # o orientador principal corresponde; não desloca nomes para outras colunas.
        hist_advisor, hist_coadvisor = split_advisors(defense.advisor_raw_excel)
        if hist_coadvisor and not parsed.coadvisor and same_person(hist_advisor, parsed.advisor):
            parsed.coadvisor = hist_coadvisor
            warnings.append("Coorientador ausente no site; recuperado do histórico.")
        elif hist_coadvisor and parsed.coadvisor and not same_person(hist_coadvisor, parsed.coadvisor):
            warnings.append("Coorientador no site difere do histórico; revisar.")
        all_members = parsed.titulares + parsed.suplentes
        advisor_aff = affiliation_for_person(parsed.advisor, all_members)
        coadvisor_aff = affiliation_for_person(parsed.coadvisor, all_members)

        # Em uma defesa do IC, o orientador principal normalmente é do IC. Só usamos esse
        # fallback para o principal; coorientador externo fica em branco se não aparecer na banca.
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
            *catalog.audit_values(),
        ])
        if delay > 0:
            time.sleep(delay)

    style_output(ws)
    audit.freeze_panes = "A2"
    audit.auto_filter.ref = f"A1:P{audit.max_row}"
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
    for col in "NOP":
        audit.column_dimensions[col].width = 55

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    logger.info("Arquivo salvo em: %s", output_path.resolve())


# ---------- atualização protegida de uma planilha já revisada ----------
# O XLSX é um pacote ZIP/OOXML. Neste modo não regravamos o workbook com
# openpyxl: alteramos somente os elementos <c> autorizados e acrescentamos
# strings à tabela de textos, preservando desenhos, estilos e metadados.
from zipfile import ZipFile
from copy import copy as copy_zipinfo
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape as xml_escape
import posixpath
import tempfile
import os

_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
_REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
_ROW = re.compile(rb'<row\b[^>]*?\br="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.S)
_CELL = re.compile(rb'<c\b[^>]*?\br="([A-Z]+[0-9]+)"[^>]*?(?:/>|>.*?</c>)', re.S)


def _colnum(ref):
    n = 0
    for char in re.match(r'[A-Z]+', ref).group():
        n = n * 26 + ord(char) - 64
    return n


def _sheet_paths(z):
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    targets = {r.attrib['Id']: r.attrib['Target'] for r in rels}
    root = ET.fromstring(z.read('xl/workbook.xml'))
    return {s.attrib['name']: (targets[s.attrib[_REL+'id']].lstrip('/') if targets[s.attrib[_REL+'id']].startswith('/') else posixpath.normpath('xl/'+targets[s.attrib[_REL+'id']]))
            for s in root.find(_NS+'sheets')}


def _read_xml_cells(raw, strings):
    root = ET.fromstring(raw)
    cells = {}
    for c in root.findall(f'{_NS}sheetData/{_NS}row/{_NS}c'):
        ref = c.attrib['r']
        if ref in cells:
            raise ValueError(f'Célula duplicada: {ref}')
        v = c.find(_NS+'v')
        if c.get('t') == 's' and v is not None:
            value = strings[int(v.text)]
        elif c.get('t') == 'inlineStr':
            value = ''.join(t.text or '' for t in c.iter(_NS+'t'))
        else:
            value = v.text if v is not None else ''
        cells[ref] = (value or '', c)
    return cells


def _patch_sheet(raw, updates, strings, string_ids, new_strings):
    """Substitui apenas células autorizadas, sem serializar novamente a folha."""
    by_row = defaultdict(dict)
    for ref, val in updates.items():
        by_row[int(re.search(r'\d+$', ref).group())][ref] = val
    applied = set()
    delta = 0
    def row_replace(match):
        nonlocal delta
        row_no = int(match[1]); original = match[0]
        changes = by_row.get(row_no)
        if not changes:
            return original
        existing = {m[1].decode(): m for m in _CELL.finditer(original)}
        edits = []
        for ref, val in sorted(changes.items(), key=lambda x: _colnum(x[0])):
            old = existing.get(ref)
            if old:
                chunk = old[0]
                elem = ET.fromstring(chunk)
                if elem.find('f') is not None or elem.find(_NS+'f') is not None:
                    raise ValueError(f'{ref} contém fórmula; atualização cancelada.')
                start = chunk.split(b'>', 1)[0].rstrip(b'/')
                start = re.sub(rb'\s+t="[^"]*"', b'', start)
                delta -= int(elem.get('t') == 's')
            else:
                # A célula antes inexistente recebe o estilo do nome no mesmo bloco
                # se disponível; nenhuma coluna ou linha é inserida.
                start = f'<c r="{ref}"'.encode()
                if val:
                    prev = [m for key,m in existing.items() if _colnum(key) < _colnum(ref)]
                    if prev:
                        style = re.search(rb'\bs="(\d+)"', prev[-1][0].split(b'>',1)[0])
                        if style: start += b' s="'+style[1]+b'"'
            if val:
                if len(val.encode('utf-16-le'))//2 > 32767:
                    raise ValueError(f'Texto de {ref} excede o limite do Excel; nada foi salvo.')
                if val not in string_ids:
                    string_ids[val] = len(strings)+len(new_strings)
                    new_strings.append(val)
                replacement = start + b' t="s"><v>'+str(string_ids[val]).encode()+b'</v></c>'
                delta += 1
            else:
                replacement = start + b'/>' if old else b''
            if old:
                edits.append((old.start(), old.end(), replacement))
            elif replacement:
                following = [m for key,m in existing.items() if _colnum(key)>_colnum(ref)]
                pos = following[0].start() if following else original.rfind(b'</row>')
                edits.append((pos,pos,replacement))
            applied.add(ref)
        for a,b,repl in reversed(sorted(edits,key=lambda e:(e[0],e[1]))):
            original = original[:a]+repl+original[b:]
        return original
    result = _ROW.sub(row_replace, raw)
    if applied != set(updates):
        raise ValueError(f'Linhas não reconhecidas: {set(updates)-applied}')
    # Prova de preservação: retirando somente as células autorizadas, o XML
    # inteiro deve ser byte a byte igual ao original.
    def mask(data):
        return _CELL.sub(lambda m: b'' if m[1].decode() in updates else m[0],data)
    if mask(raw) != mask(result):
        raise AssertionError('Alteração fora das células autorizadas; saída cancelada.')
    return result, delta


def refresh_existing(input_path: Path, output_path: Path, year: int, catalog: PrizeCatalog,
                     logger: logging.Logger) -> None:
    """Modo protegido: só Prêmio e J:M da Auditoria; nenhuma consulta à rede."""
    if output_path.exists():
        raise FileExistsError(f'A saída já existe: {output_path}. Escolha outro nome para não sobrescrever arquivos.')
    with ZipFile(input_path) as z:
        paths = _sheet_paths(z)
        title = f'Classificação {year}'
        if title not in paths or 'Auditoria' not in paths:
            raise ValueError(f'Esperadas abas {title!r} e Auditoria.')
        sp, ap = paths[title], paths['Auditoria']
        sst_path = 'xl/sharedStrings.xml'
        # Formato dos três arquivos revisados enviados: shared strings e XML padrão.
        if sst_path not in z.namelist():
            raise ValueError('Formato XLSX sem tabela de strings compartilhadas: modo protegido não aplicável. Nenhum arquivo alterado.')
        sst_raw = z.read(sst_path)
        sst = ET.fromstring(sst_raw)
        strings = [''.join(t.text or '' for t in si.iter(_NS+'t')) for si in sst.findall(_NS+'si')]
        ids = {s:i for i,s in enumerate(strings)}
        data_raw, audit_raw = z.read(sp), z.read(ap)
        data, audit = _read_xml_cells(data_raw,strings), _read_xml_cells(audit_raw,strings)
        def val(cells, ref): return cells.get(ref,('',None))[0]
        headers = [val(data,get_column_letter(c)+'1') for c in range(1,53)]
        if headers[:2] != OUTPUT_HEADERS[:2] or any(
            headers[i:i+3] != OUTPUT_HEADERS[i:i+3] or set(headers[i+3:i+5]) != {'Gênero','Raça'}
            for i in range(2,52,5)):
            raise ValueError('Estrutura das 52 colunas não reconhecida. Nenhum arquivo alterado.')
        audit_headers = ['Linha histórico','Classificação','Aluno','Data','URL','Status página','Orientação (Excel)','Orientação (site)','Avisos','Matches de prêmio','Edições de prêmio disponíveis','Edições confirmadas','Edições não confirmadas']
        if [val(audit,get_column_letter(c)+'1') for c in range(1,14)] != audit_headers:
            raise ValueError('Cabeçalho de Auditoria não reconhecido.')
        rows = sorted(int(k[1:]) for k,v in data.items() if re.fullmatch(r'B\d+',k) and int(k[1:])>1 and v[0].strip())
        ars = sorted(int(k[1:]) for k,v in audit.items() if re.fullmatch(r'C\d+',k) and int(k[1:])>1 and v[0].strip())
        def key(cells,r,cols): return tuple(clean_space(val(cells,c+str(r))) for c in cols)
        amap = {key(audit,r,'BC'):r for r in ars}
        keys = [key(data,r,'AB') for r in rows]
        if len(amap)!=len(ars) or len(set(keys))!=len(keys) or set(keys)!=set(amap):
            raise ValueError('Defesas e Auditoria não correspondem univocamente por classificação e aluno.')
        updates, aud_updates = {}, {}
        modes=defaultdict(int); coadvisors=0; named=0; orphan=0
        for r,k in zip(rows,keys):
            notes=[]; ar=amap[k]
            for c in range(3,53,5):
                name_ref, aff_ref, prize_ref = [get_column_letter(x)+str(r) for x in (c,c+1,c+2)]
                name, aff = clean_space(val(data,name_ref)), clean_space(val(data,aff_ref))
                old=val(data,prize_ref)
                prize, mode, score = prize_text_for_person(name,aff,catalog,logger)
                if name:
                    named+=1; modes[mode.split(':')[0]]+=1
                    coadvisors+=int(c==8)
                    notes.append(f'{headers[c-1]}: {mode} ({score:.3f})')
                elif old:
                    orphan+=1
                    notes.append(f'{headers[c-1]}: prêmio antigo sem membro; removido de {prize_ref}')
                if old != prize: updates[prize_ref]=prize
            notes += [f'Fonte local: {catalog.path.name}; SHA256: {catalog.sha256}', catalog.criteria, catalog.coverage]
            for c,new in zip('JKLM', [' | '.join(notes), *catalog.audit_values()[:3]]):
                ref=c+str(ar)
                if val(audit,ref)!=new: aud_updates[ref]=new
        new_strings=[]
        d2,delta1 = _patch_sheet(data_raw,updates,strings,ids,new_strings)
        a2,delta2 = _patch_sheet(audit_raw,aud_updates,strings,ids,new_strings)
        # Adiciona textos no final; não modifica qualquer string original.
        extra=b''.join(('<si><t xml:space="preserve">'+xml_escape(s)+'</t></si>').encode('utf-8') for s in new_strings)
        if not sst_raw.endswith(b'</sst>'):
            raise ValueError('Tabela de strings em formato inesperado.')
        s2=sst_raw[:-6]+extra+b'</sst>'
        if 'uniqueCount' in sst.attrib:
            s2=re.sub(rb'\buniqueCount="\d+"',f'uniqueCount="{len(strings)+len(new_strings)}"'.encode(),s2,count=1)
        if 'count' in sst.attrib:
            s2=re.sub(rb'\bcount="\d+"',f'count="{int(sst.attrib["count"])+delta1+delta2}"'.encode(),s2,count=1)
        # Segunda verificação, com leitura lógica de todas as células.
        all_strings=strings+new_strings
        for before,after,allowed in [(data,_read_xml_cells(d2,all_strings),updates),(audit,_read_xml_cells(a2,all_strings),aud_updates)]:
            for ref in set(before)|set(after):
                if ref in allowed:
                    if val(after,ref)!=allowed[ref]: raise AssertionError(f'Valor incorreto: {ref}')
                elif val(before,ref)!=val(after,ref):
                    raise AssertionError(f'Valor protegido modificado: {ref}')
        changes={sp:d2,ap:a2,sst_path:s2}
        output_path.parent.mkdir(parents=True,exist_ok=True)
        fd,tmp=tempfile.mkstemp(suffix='.xlsx',dir=output_path.parent); os.close(fd)
        try:
            with ZipFile(tmp,'w') as out:
                out.comment=z.comment
                for item in z.infolist():
                    out.writestr(copy_zipinfo(item), changes.get(item.filename,z.read(item.filename)))
            with ZipFile(tmp) as out:
                if out.testzip() is not None or out.namelist()!=z.namelist():
                    raise AssertionError('Pacote XLSX inválido.')
                for filename in z.namelist():
                    expected=changes.get(filename,z.read(filename))
                    if out.read(filename)!=expected:
                        raise AssertionError(f'Parte do arquivo alterada indevidamente: {filename}')
            # Não sobrescreve nem uma saída criada por outro processo.
            with open(output_path,'xb') as out, open(tmp,'rb') as inp:
                import shutil
                shutil.copyfileobj(inp,out)
        finally:
            os.unlink(tmp)
    logger.info('VALIDADO: %d defesas | %d coorientações preservadas | %d membros | %d células Prêmio modificadas | %d prêmios órfãos limpos | %d células Auditoria modificadas',len(rows),coadvisors,named,len(updates),orphan,len(aud_updates))
    logger.info('Todas as outras células, estilos, desenhos e partes do pacote foram preservados. %s',output_path)
    logger.info('Resultados de correspondência: %s',dict(modes))


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Monta a classificação das defesas do IC/UNICAMP e cruza com a base validada local de prêmios do IC.")
    p.add_argument("--gerar-do-zero", action="store_true", help="Autoriza explicitamente a reextração de novas bancas; exige revisão humana.")
    p.add_argument("--historico", type=Path, help="XLSX com o histórico de defesas (necessário para gerar do zero).")
    p.add_argument("--atualizar-classificacao", type=Path, help="XLSX já preenchido: atualiza apenas prêmios e auditoria, sem rede.")
    p.add_argument("--premios-base", type=Path, default=DEFAULT_PRIZE_BASE, help="CSV final validado de prêmios do IC.")
    p.add_argument("--premios-ate-defesa", action="store_true", help="Limita os prêmios até o ano da defesa, inclusive.")
    p.add_argument("--premios-somente-conquistas", action="store_true", help="Inclui somente premiações/colocações e menções honrosas.")
    p.add_argument("--template", type=Path, default=None, help="XLSX Classificação.xlsx (opcional, preserva o formato do cabeçalho).")
    p.add_argument("--saida", type=Path, default=Path("Classificação_2023_preenchida.xlsx"), help="Arquivo XLSX de saída.")
    p.add_argument("--ano", type=int, default=2023, help="Ano das defesas (default: 2023).")
    p.add_argument(
        "--premios-anos",
        default="auto",
        help=(
            "Anos do CSV local a considerar. Default: auto (todos os anos presentes na base). "
            "Também aceita faixa, ex.: 2019-2025, ou lista, ex.: 2019,2021,2024."
        ),
    )
    # Compatibilidade com a versão anterior do script: esses argumentos antigos
    # são aceitos para não quebrar comandos já copiados, mas não limitam a busca.
    p.add_argument("--premios-ano", type=int, default=None, help=argparse.SUPPRESS)
    p.add_argument("--premios-todos", action="store_true", help=argparse.SUPPRESS)
    p.add_argument("--delay", type=float, default=0.25, help="Pausa entre defesas, em segundos (default: 0.25).")
    p.add_argument(
        "--cache-dir",
        type=Path,
        default=Path(".cache_premios_unicamp"),
        help=argparse.SUPPRESS,
    )
    p.add_argument(
        "--refresh-cache",
        action="store_true",
        help=argparse.SUPPRESS,
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

    if args.premios_ano is not None or args.premios_todos or args.refresh_cache or args.cache_dir != Path(".cache_premios_unicamp"):
        logger.warning("Argumentos antigos de consulta/cache do Anuário são ignorados; use --premios-base e --premios-anos.")
    if bool(args.historico) == bool(args.atualizar_classificacao):
        logger.error("Informe exatamente um modo: --historico ou --atualizar-classificacao.")
        return 2
    if args.historico and not args.gerar_do_zero:
        logger.error("Para preservar planilhas revisadas, use --atualizar-classificacao. Para uma NOVA planilha, informe --gerar-do-zero junto com --historico.")
        return 2
    if args.gerar_do_zero and args.atualizar_classificacao:
        logger.error("Não combine --gerar-do-zero com --atualizar-classificacao.")
        return 2
    if args.saida.exists():
        logger.error("Arquivo de saída já existe; escolha outro nome: %s", args.saida)
        return 2
    inputs = [x for x in (args.historico, args.template, args.atualizar_classificacao, args.premios_base) if x is not None]
    if any(not x.is_file() for x in inputs):
        logger.error("Arquivo de entrada não encontrado: %s", [str(x) for x in inputs if not x.is_file()])
        return 2
    if args.saida.resolve() in {x.resolve() for x in inputs}:
        logger.error("A saída não pode sobrescrever um arquivo de entrada.")
        return 2
    try:
        if args.atualizar_classificacao:
            catalog = PrizeCatalog(args.premios_base, args.premios_anos, args.ano,
                                   args.premios_ate_defesa, args.premios_somente_conquistas, logger)
            refresh_existing(args.atualizar_classificacao, args.saida, args.ano, catalog, logger)
        else:
            process(history_path=args.historico, template_path=args.template, output_path=args.saida,
                    year=args.ano, prizes_year_spec=args.premios_anos, delay=max(0.0, args.delay),
                    logger=logger, prizes_base=args.premios_base, until_defense=args.premios_ate_defesa,
                    only_achievements=args.premios_somente_conquistas)
    except KeyboardInterrupt:
        logger.error("Execução interrompida pelo usuário.")
        return 130
    except Exception as exc:
        logger.exception("Falha fatal: %s", exc)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
