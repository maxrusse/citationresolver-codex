#!/usr/bin/env python3
"""Convert plain numeric citations in a .docx into live Zotero field citations.

Features:
- Detect in-text numeric citations like (1), (2, 4), (13-18), (13–18)
- Parse numbered references from the "B1. References" section
- Match references against local Zotero library (`zotero.sqlite`) when available
- Optionally add unmatched references to local Zotero via connector HTTP endpoint
- Write Word field runs with ADDIN ZOTERO_ITEM CSL_CITATION JSON code
"""

from __future__ import annotations

import argparse
import copy
import difflib
import html
import io
import json
import os
import random
import re
import shutil
import sqlite3
import string
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from urllib import error, request
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
CP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
W = f"{{{W_NS}}}"

CITE_PAREN_RE = re.compile(r"\((\d+(?:\s*[-–—,]\s*\d+)*)\)")
CITE_BRACKET_RE = re.compile(r"\[(\d+(?:\s*[-–—,]\s*\d+)*)\]")
CITE_SUPER_TOKEN_RE = re.compile(r"\(?\[?(\d+(?:\s*[-–—,]\s*\d+)*)\]?\)?")
REF_LINE_RE = re.compile(r"^\s*(\d+)\.\s+(.+?)\s*$")
DOI_RE = re.compile(r"10\.\d{4,9}/[-._;()/:A-Z0-9]+", re.IGNORECASE)
YEAR_RE = re.compile(r"\b(19\d{2}|20\d{2}|21\d{2})\b")


@dataclass
class RefEntry:
    number: int
    raw: str
    title: str | None
    doi: str | None
    year: int | None
    authors: list[dict[str, str]]


@dataclass
class ZoteroItem:
    item_id: int
    key: str
    library_id: int
    title: str | None
    doi: str | None
    year: int | None
    first_author: str | None


@dataclass
class MatchResult:
    item: ZoteroItem | None
    method: str | None
    score: float


def _rand_id(n: int = 8) -> str:
    return "".join(random.choices(string.ascii_letters + string.digits, k=n))


def _norm_text(value: str | None) -> str:
    if not value:
        return ""
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", value.lower())).strip()


def _clean_doi(value: str | None) -> str | None:
    if not value:
        return None
    s = value.strip()
    s = s.replace("\u200b", "")
    s = s.strip("\"'`“”‘’")
    s = s.replace("https://doi.org/", "").replace("http://doi.org/", "").replace("doi:", "").strip()
    s = s.strip("()[]{}")
    while s and s[-1] in ".,;:)]}\"'":
        s = s[:-1]
    while s and s[0] in "\"'([{":
        s = s[1:]
    s = s.strip()
    return s or None


def _norm_doi(value: str | None) -> str:
    cleaned = _clean_doi(value)
    if not cleaned:
        return ""
    s = cleaned.lower()
    return s


def _parse_year(value: str | None) -> int | None:
    if not value:
        return None
    m = YEAR_RE.search(value)
    return int(m.group(1)) if m else None


def _get_run_text(run: ET.Element) -> str:
    texts = [t.text or "" for t in run.findall(f"{W}t")]
    return "".join(texts)


def _parse_cite_token(token: str) -> list[int]:
    nums: list[int] = []
    for part in [p.strip() for p in token.split(",")]:
        rng = re.split(r"\s*[-–—]\s*", part)
        if len(rng) == 2 and rng[0].isdigit() and rng[1].isdigit():
            a, b = int(rng[0]), int(rng[1])
            if a <= b:
                nums.extend(range(a, b + 1))
            else:
                nums.extend(range(b, a + 1))
        elif part.isdigit():
            nums.append(int(part))
    out: list[int] = []
    seen: set[int] = set()
    for n in nums:
        if n not in seen:
            out.append(n)
            seen.add(n)
    return out


def _parse_authors(author_segment: str) -> list[dict[str, str]]:
    cleaned = author_segment.strip().rstrip(".")
    if not cleaned:
        return []
    cleaned = re.sub(r"\bet al\.?\b", "", cleaned, flags=re.IGNORECASE).strip().rstrip(",")
    parts = [p.strip() for p in cleaned.split(",") if p.strip()]
    authors: list[dict[str, str]] = []
    for p in parts[:8]:
        toks = p.split()
        if not toks:
            continue
        family = toks[0]
        given = " ".join(toks[1:]) if len(toks) > 1 else ""
        authors.append({"family": family, "given": given})
    return authors


def _parse_reference_entry(number: int, raw: str) -> RefEntry:
    m_doi = DOI_RE.search(raw)
    doi = _clean_doi(m_doi.group(0)) if m_doi else None
    m_year = YEAR_RE.search(raw)
    year = int(m_year.group(1)) if m_year else None

    parts = [p.strip() for p in raw.split(". ") if p.strip()]
    author_seg = parts[0] if parts else ""
    title = parts[1] if len(parts) > 1 else None
    authors = _parse_authors(author_seg)
    return RefEntry(number=number, raw=raw, title=title, doi=doi, year=year, authors=authors)


def _make_item_data(ref: RefEntry, num: int) -> dict:
    item = {
        "id": f"ref-{num}",
        "type": "article-journal",
        "title": ref.title or ref.raw[:180],
    }
    if ref.authors:
        item["author"] = ref.authors
    if ref.year:
        item["issued"] = {"date-parts": [[ref.year]]}
    if ref.doi:
        item["DOI"] = ref.doi
    return item


class LocalZoteroIndex:
    def __init__(self, db_path: Path, local_user_key: str | None = None):
        self.db_path = db_path
        self.local_user_key = local_user_key
        self.items: list[ZoteroItem] = []
        self.by_doi: dict[str, list[ZoteroItem]] = {}

    def load(self) -> None:
        uri = f"file:{self.db_path.as_posix()}?immutable=1"
        con = sqlite3.connect(uri, uri=True)
        cur = con.cursor()

        if not self.local_user_key:
            cur.execute("SELECT value FROM settings WHERE setting='account' AND key='localUserKey'")
            row = cur.fetchone()
            self.local_user_key = str(row[0]) if row and row[0] else "auto"

        cur.execute(
            """
            SELECT
              i.itemID,
              i.key,
              i.libraryID,
              MAX(CASE WHEN f.fieldName='title' THEN v.value END) AS title,
              MAX(CASE WHEN f.fieldName='DOI' THEN v.value END) AS doi,
              MAX(CASE WHEN f.fieldName='date' THEN v.value END) AS dateval
            FROM items i
            LEFT JOIN deletedItems di ON di.itemID = i.itemID
            LEFT JOIN itemAttachments ia ON ia.itemID = i.itemID
            LEFT JOIN itemNotes ino ON ino.itemID = i.itemID
            LEFT JOIN itemData d ON d.itemID = i.itemID
            LEFT JOIN fields f ON f.fieldID = d.fieldID
            LEFT JOIN itemDataValues v ON v.valueID = d.valueID
            WHERE di.itemID IS NULL
              AND ia.itemID IS NULL
              AND ino.itemID IS NULL
            GROUP BY i.itemID, i.key, i.libraryID
            HAVING
              MAX(CASE WHEN f.fieldName='title' THEN v.value END) IS NOT NULL
              OR MAX(CASE WHEN f.fieldName='DOI' THEN v.value END) IS NOT NULL
            """
        )
        rows = cur.fetchall()

        first_author_by_item: dict[int, str] = {}
        cur.execute(
            """
            SELECT ic.itemID, c.lastName, ic.orderIndex
            FROM itemCreators ic
            JOIN creators c ON c.creatorID = ic.creatorID
            ORDER BY ic.itemID, ic.orderIndex
            """
        )
        for item_id, last_name, _order_idx in cur.fetchall():
            if item_id not in first_author_by_item and last_name:
                first_author_by_item[item_id] = last_name

        items: list[ZoteroItem] = []
        for item_id, key, library_id, title, doi, dateval in rows:
            zi = ZoteroItem(
                item_id=int(item_id),
                key=str(key),
                library_id=int(library_id),
                title=str(title) if title else None,
                doi=str(doi) if doi else None,
                year=_parse_year(str(dateval) if dateval else None),
                first_author=first_author_by_item.get(int(item_id)),
            )
            items.append(zi)

        self.items = items
        by_doi: dict[str, list[ZoteroItem]] = {}
        for item in items:
            d = _norm_doi(item.doi)
            if not d:
                continue
            by_doi.setdefault(d, []).append(item)
        self.by_doi = by_doi
        con.close()

    def uri_for(self, item: ZoteroItem) -> str:
        return f"http://zotero.org/users/local/{self.local_user_key}/items/{item.key}"

    def match(self, ref: RefEntry) -> MatchResult:
        d = _norm_doi(ref.doi)
        if d and d in self.by_doi and self.by_doi[d]:
            cands = self.by_doi[d]
            # Prefer same year if possible.
            if ref.year:
                year_match = [c for c in cands if c.year == ref.year]
                if year_match:
                    return MatchResult(item=year_match[0], method="doi+year", score=1.0)
            return MatchResult(item=cands[0], method="doi", score=0.99)

        ref_title = _norm_text(ref.title or ref.raw)
        if not ref_title:
            return MatchResult(item=None, method=None, score=0.0)
        ref_author = _norm_text(ref.authors[0]["family"] if ref.authors else "")
        best: tuple[ZoteroItem | None, float] = (None, 0.0)
        for item in self.items:
            item_title = _norm_text(item.title)
            if not item_title:
                continue
            ratio = difflib.SequenceMatcher(a=ref_title, b=item_title).ratio()
            if ratio < 0.84:
                continue
            score = ratio
            if ref.year and item.year and ref.year == item.year:
                score += 0.05
            if ref_author and item.first_author and _norm_text(item.first_author) == ref_author:
                score += 0.03
            if score > best[1]:
                best = (item, score)

        if best[0] and best[1] >= 0.90:
            return MatchResult(item=best[0], method="title_fuzzy", score=float(best[1]))
        return MatchResult(item=None, method=None, score=float(best[1]))


def _http_post_json(url: str, payload: dict, timeout: int = 10) -> tuple[int, str]:
    data = json.dumps(payload).encode("utf-8")
    req = request.Request(url, data=data, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("Accept", "application/json")
    with request.urlopen(req, timeout=timeout) as resp:
        body = resp.read().decode("utf-8", errors="replace")
        return int(resp.status), body


def _detect_zotero_version(connector_base: str) -> str:
    url = connector_base.rstrip("/") + "/connector/ping"
    req = request.Request(url, method="GET")
    try:
        with request.urlopen(req, timeout=5) as resp:
            header = resp.headers.get("X-Zotero-Version")
            if header:
                return str(header).strip()
    except Exception:  # noqa: BLE001
        pass
    return "7.0.0"


def _build_zotero_pref_value(session_id: str, style_id: str, zotero_version: str) -> str:
    return (
        f'<data data-version="3" zotero-version="{zotero_version}">'
        f'<session id="{session_id}"/>'
        f'<style id="{style_id}" hasBibliography="1" bibliographyStyleHasBeenSet="1"/>'
        '<prefs><pref name="fieldType" value="Field"/></prefs>'
        "</data>"
    )


def _update_custom_props_xml(existing_bytes: bytes | None, pref_value: str) -> bytes:
    ET.register_namespace("", CP_NS)
    ET.register_namespace("vt", VT_NS)

    if existing_bytes:
        root = ET.fromstring(existing_bytes)
    else:
        root = ET.Element(f"{{{CP_NS}}}Properties")

    existing_prop = None
    max_pid = 1
    for prop in root.findall(f"{{{CP_NS}}}property"):
        pid = prop.get("pid")
        if pid and pid.isdigit():
            max_pid = max(max_pid, int(pid))
        if prop.get("name", "").startswith("ZOTERO_PREF_"):
            existing_prop = prop

    if existing_prop is None:
        prop = ET.SubElement(root, f"{{{CP_NS}}}property")
        prop.set("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        prop.set("pid", str(max_pid + 1))
        prop.set("name", "ZOTERO_PREF_1")
    else:
        prop = existing_prop
        for child in list(prop):
            prop.remove(child)

    lp = ET.SubElement(prop, f"{{{VT_NS}}}lpwstr")
    lp.text = pref_value
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _extract_style_id_from_custom_props(custom_xml_bytes: bytes | None) -> str | None:
    if not custom_xml_bytes:
        return None
    try:
        root = ET.fromstring(custom_xml_bytes)
    except ET.ParseError:
        return None

    # Namespace-aware lookup first
    props = root.findall(f"{{{CP_NS}}}property")
    # Fallback for malformed/namespace-light XML
    if not props:
        props = root.findall(".//property")

    for prop in props:
        if not prop.get("name", "").startswith("ZOTERO_PREF_"):
            continue
        lp = prop.find(f".//{{{VT_NS}}}lpwstr")
        if lp is None:
            lp = prop.find(".//lpwstr")
        if lp is None or not lp.text:
            continue
        raw = html.unescape(lp.text)
        try:
            pref_root = ET.fromstring(raw)
        except ET.ParseError:
            continue
        style_node = pref_root.find("style")
        if style_node is not None and style_node.get("id"):
            return style_node.get("id")
    return None


def _add_missing_refs_to_zotero(
    refs: list[RefEntry],
    connector_base: str,
) -> dict:
    if not refs:
        return {"attempted": 0, "added": 0, "errors": []}

    items = []
    for ref in refs:
        creators = []
        for a in ref.authors[:8]:
            creators.append(
                {
                    "creatorType": "author",
                    "firstName": a.get("given", ""),
                    "lastName": a.get("family", ""),
                }
            )
        item = {
            "id": f"ref-{ref.number}",
            "itemType": "journalArticle",
            "title": ref.title or ref.raw[:250],
            "creators": creators,
            "DOI": ref.doi or "",
            "date": str(ref.year) if ref.year else "",
            "extra": f"Imported by docx_zotero_integrator from numbered reference {ref.number}",
        }
        items.append(item)

    payload = {"sessionID": f"docx-{_rand_id(12)}", "items": items}
    url = connector_base.rstrip("/") + "/connector/saveItems"
    errors_out: list[str] = []
    try:
        status, body = _http_post_json(url, payload, timeout=20)
        if status not in (200, 201):
            errors_out.append(f"saveItems status {status}: {body[:500]}")
            return {"attempted": len(refs), "added": 0, "errors": errors_out}
    except error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        errors_out.append(f"HTTP {e.code}: {body[:500]}")
        return {"attempted": len(refs), "added": 0, "errors": errors_out}
    except Exception as e:  # noqa: BLE001
        errors_out.append(str(e))
        return {"attempted": len(refs), "added": 0, "errors": errors_out}
    return {"attempted": len(refs), "added": len(refs), "errors": []}


def _copy_rpr(run: ET.Element) -> ET.Element | None:
    rpr = run.find(f"{W}rPr")
    return copy.deepcopy(rpr) if rpr is not None else None


def _mk_run_with_text(orig_run: ET.Element, text: str) -> ET.Element:
    r = ET.Element(f"{W}r", attrib=dict(orig_run.attrib))
    rpr = _copy_rpr(orig_run)
    if rpr is not None:
        r.append(rpr)
    t = ET.SubElement(r, f"{W}t")
    if text.startswith(" ") or text.endswith(" "):
        t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = text
    return r


def _mk_fldchar_run(orig_run: ET.Element, fld_type: str) -> ET.Element:
    r = ET.Element(f"{W}r", attrib=dict(orig_run.attrib))
    rpr = _copy_rpr(orig_run)
    if rpr is not None:
        r.append(rpr)
    fld = ET.SubElement(r, f"{W}fldChar")
    fld.set(f"{W}fldCharType", fld_type)
    return r


def _mk_instr_run(orig_run: ET.Element, instruction: str) -> ET.Element:
    r = ET.Element(f"{W}r", attrib=dict(orig_run.attrib))
    rpr = _copy_rpr(orig_run)
    if rpr is not None:
        r.append(rpr)
    instr = ET.SubElement(r, f"{W}instrText")
    instr.set(f"{{{XML_NS}}}space", "preserve")
    instr.text = instruction
    return r


def _mk_empty_run() -> ET.Element:
    return ET.Element(f"{W}r")


def _mk_bibliography_instruction() -> str:
    payload = {"uncited": [], "omitted": [], "custom": []}
    return " ADDIN ZOTERO_BIBL " + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + " CSL_BIBLIOGRAPHY "


def _first_run_in_paragraph(para: ET.Element) -> ET.Element:
    for c in list(para):
        if c.tag == f"{W}r":
            return c
    return _mk_empty_run()


def _create_bibliography_paragraph(template_para: ET.Element | None) -> ET.Element:
    p = ET.Element(f"{W}p")
    if template_para is not None:
        ppr = template_para.find(f"{W}pPr")
        if ppr is not None:
            p.append(copy.deepcopy(ppr))
    template_run = _first_run_in_paragraph(template_para) if template_para is not None else _mk_empty_run()
    p.append(_mk_fldchar_run(template_run, "begin"))
    p.append(_mk_instr_run(template_run, _mk_bibliography_instruction()))
    p.append(_mk_fldchar_run(template_run, "separate"))
    p.append(_mk_run_with_text(template_run, "Bibliography will be generated by Zotero on refresh."))
    p.append(_mk_fldchar_run(template_run, "end"))
    return p


def _find_reference_block_end(paragraphs: list[ET.Element], start_idx: int) -> int:
    # End is exclusive.
    i = start_idx + 1
    while i < len(paragraphs):
        txt = _extract_paragraph_text(paragraphs[i]).strip()
        if not txt:
            i += 1
            continue
        if REF_LINE_RE.match(txt):
            i += 1
            continue
        break
    return i


def _insert_or_replace_bibliography_field(
    root: ET.Element,
    paragraphs: list[ET.Element],
    ref_start: int,
    replace_reference_list: bool,
) -> bool:
    # Skip if bibliography field already exists.
    for p in paragraphs:
        for instr in p.findall(f".//{W}instrText"):
            if "ZOTERO_BIBL" in (instr.text or ""):
                return False

    body = root.find(f".//{W}body")
    if body is None:
        return False
    body_children = list(body)
    p_to_idx = {id(p): idx for idx, p in enumerate(body_children) if p.tag == f"{W}p"}

    start_para = paragraphs[ref_start]
    heading_body_idx = p_to_idx.get(id(start_para))
    if heading_body_idx is None:
        return False

    block_end = _find_reference_block_end(paragraphs, ref_start)
    first_ref_para = paragraphs[ref_start + 1] if (ref_start + 1) < len(paragraphs) else None
    new_bibl_para = _create_bibliography_paragraph(first_ref_para)

    insert_at = heading_body_idx + 1
    if replace_reference_list and first_ref_para is not None:
        # Rebuild paragraph -> body index map after possible earlier edits.
        body_children = list(body)
        p_to_idx = {id(p): idx for idx, p in enumerate(body_children) if p.tag == f"{W}p"}
        start_ref_idx = p_to_idx.get(id(first_ref_para))
        if start_ref_idx is not None:
            # Collect paragraphs in the current reference list block.
            refs_to_remove: list[ET.Element] = []
            for p in paragraphs[ref_start + 1 : block_end]:
                refs_to_remove.append(p)
            for p in refs_to_remove:
                if p in body:
                    body.remove(p)
            insert_at = start_ref_idx

    body.insert(insert_at, new_bibl_para)
    return True


def _build_citation_instruction(
    citation_literal: str,
    token: str,
    refs: dict[int, RefEntry],
    matches: dict[int, MatchResult],
    index: LocalZoteroIndex | None,
) -> str:
    nums = _parse_cite_token(token)
    items = []
    for n in nums:
        ref = refs.get(
            n,
            RefEntry(number=n, raw=f"Reference {n}", title=f"Reference {n}", doi=None, year=None, authors=[]),
        )
        m = matches.get(n)
        if m and m.item and index:
            cid = m.item.item_id
            uris = [index.uri_for(m.item)]
        else:
            cid = f"ref-{n}"
            uris = [f"http://zotero.org/users/local/auto/items/REF{n:04d}"]
        items.append(
            {
                "id": cid,
                "uris": uris,
                "itemData": _make_item_data(ref, n),
            }
        )

    payload = {
        "citationID": _rand_id(),
        "properties": {
            "formattedCitation": citation_literal,
            # Keep plainCitation aligned with visible text to avoid Zotero
            # "citation changed manually" prompts on first edit.
            "plainCitation": citation_literal,
            "noteIndex": 0,
        },
        "citationItems": items,
        "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json",
    }
    return " ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(payload, ensure_ascii=False) + " "


def _extract_paragraph_text(para: ET.Element) -> str:
    return "".join(t.text or "" for t in para.findall(f".//{W}t"))


def _find_references_start(paragraphs: list[ET.Element]) -> int | None:
    for i, p in enumerate(paragraphs):
        txt = _extract_paragraph_text(p).strip()
        if txt.lower().startswith("b1. references"):
            return i
    return None


def _parse_reference_map(paragraphs: Iterable[ET.Element], start_idx: int) -> dict[int, RefEntry]:
    refs: dict[int, RefEntry] = {}
    for p in list(paragraphs)[start_idx + 1 :]:
        txt = _extract_paragraph_text(p).strip()
        if not txt:
            continue
        m = REF_LINE_RE.match(txt)
        if not m:
            continue
        n = int(m.group(1))
        refs[n] = _parse_reference_entry(n, m.group(2))
    return refs


def _run_is_superscript(run: ET.Element) -> bool:
    rpr = run.find(f"{W}rPr")
    if rpr is None:
        return False
    va = rpr.find(f"{W}vertAlign")
    if va is None:
        return False
    return va.get(f"{W}val") == "superscript"


def _find_citation_matches_in_run(run: ET.Element, run_text: str, pattern: str) -> list[dict[str, str | int]]:
    out: list[dict[str, str | int]] = []
    if not run_text:
        return out

    if pattern == "paren":
        for m in CITE_PAREN_RE.finditer(run_text):
            out.append({"start": m.start(), "end": m.end(), "literal": m.group(0), "token": m.group(1)})
        return out

    if pattern == "bracket":
        for m in CITE_BRACKET_RE.finditer(run_text):
            out.append({"start": m.start(), "end": m.end(), "literal": m.group(0), "token": m.group(1)})
        return out

    if pattern == "superscript":
        if not _run_is_superscript(run):
            return out
        stripped = run_text.strip()
        m = CITE_SUPER_TOKEN_RE.fullmatch(stripped)
        if not m:
            return out
        start = run_text.find(stripped)
        end = start + len(stripped)
        out.append({"start": start, "end": end, "literal": stripped, "token": m.group(1)})
        return out

    return out


def _detect_citation_pattern(
    paragraphs: list[ET.Element],
    ref_start: int,
    ref_numbers: set[int],
    include_superscript: bool,
) -> tuple[str, dict[str, dict[str, float | int]]]:
    candidates = ["paren", "bracket"] + (["superscript"] if include_superscript else [])
    stats: dict[str, dict[str, float | int]] = {}
    best_pattern = "paren"
    best_score = -1.0

    for cand in candidates:
        occ = 0
        total_nums = 0
        valid_nums = 0
        valid_occ = 0

        for p_idx, para in enumerate(paragraphs):
            if p_idx >= ref_start:
                break
            for child in list(para):
                if child.tag != f"{W}r":
                    continue
                if child.find(f"{W}instrText") is not None or child.find(f"{W}fldChar") is not None:
                    continue
                t_nodes = child.findall(f"{W}t")
                if len(t_nodes) != 1:
                    continue
                run_text = _get_run_text(child)
                hits = _find_citation_matches_in_run(child, run_text, cand)
                for h in hits:
                    nums = _parse_cite_token(str(h["token"]))
                    if not nums:
                        continue
                    occ += 1
                    total_nums += len(nums)
                    hit_valid = sum(1 for n in nums if n in ref_numbers)
                    valid_nums += hit_valid
                    if hit_valid > 0:
                        valid_occ += 1

        if occ == 0 or total_nums == 0:
            score = 0.0
            valid_occ_ratio = 0.0
            valid_num_ratio = 0.0
        else:
            valid_occ_ratio = valid_occ / occ
            valid_num_ratio = valid_nums / total_nums
            volume = min(occ, 20) / 20
            score = 0.55 * valid_occ_ratio + 0.35 * valid_num_ratio + 0.10 * volume

        stats[cand] = {
            "occurrences": occ,
            "numbers_total": total_nums,
            "numbers_valid": valid_nums,
            "valid_occ_ratio": round(valid_occ_ratio, 4),
            "valid_num_ratio": round(valid_num_ratio, 4),
            "score": round(score, 4),
        }
        if score > best_score:
            best_score = score
            best_pattern = cand

    return best_pattern, stats


def _inject_original_root_namespaces(serialized_xml: str, original_xml: str) -> str:
    # xml.etree may drop namespace declarations that are only referenced via
    # mc:Ignorable tokens. Re-add all original root xmlns declarations.
    m_new = re.search(r"<w:document\b[^>]*>", serialized_xml)
    m_orig = re.search(r"<w:document\b[^>]*>", original_xml)
    if not m_new or not m_orig:
        return serialized_xml

    new_tag = m_new.group(0)
    orig_tag = m_orig.group(0)
    missing_parts: list[str] = []

    # Copy prefixed namespace declarations
    for prefix, uri in re.findall(r'\sxmlns:([A-Za-z_][\w\-.]*)="([^"]+)"', orig_tag):
        token = f'xmlns:{prefix}="'
        if token not in new_tag:
            missing_parts.append(f' xmlns:{prefix}="{uri}"')

    # Copy default xmlns if present
    m_default = re.search(r'\sxmlns="([^"]+)"', orig_tag)
    if m_default and 'xmlns="' not in new_tag:
        missing_parts.append(f' xmlns="{m_default.group(1)}"')

    if not missing_parts:
        return serialized_xml

    patched_tag = new_tag[:-1] + "".join(missing_parts) + ">"
    return serialized_xml.replace(new_tag, patched_tag, 1)


def _register_namespaces_from_xml(xml_bytes: bytes) -> None:
    # Preserve all original namespace prefixes so ET doesn't emit ns0/ns1 and
    # doesn't drop declarations used by extended Word markup (w14/w15/etc.).
    seen: set[tuple[str, str]] = set()
    for _event, (prefix, uri) in ET.iterparse(io.BytesIO(xml_bytes), events=("start-ns",)):
        pfx = prefix or ""
        key = (pfx, uri)
        if key in seen:
            continue
        seen.add(key)
        try:
            ET.register_namespace(pfx, uri)
        except ValueError:
            # Ignore illegal or reserved prefixes and continue.
            pass


def _word_update_docx_fields(doc_path: Path, visible: bool = False) -> dict:
    if os.name != "nt":
        return {"attempted": False, "updated": False, "error": "Word COM update is only supported on Windows"}

    p = str(doc_path)
    p_ps = p.replace("'", "''")
    vis = "$true" if visible else "$false"
    ps_script = (
        "$ErrorActionPreference='Stop';"
        "$word=New-Object -ComObject Word.Application;"
        f"$word.Visible={vis};"
        f"$doc=$word.Documents.Open('{p_ps}');"
        "$doc.Fields.Update()|Out-Null;"
        "$range=$doc.Range();"
        "$range.Fields.Update()|Out-Null;"
        "$story=$doc.StoryRanges;"
        "while($story -ne $null){"
        "  $story.Fields.Update()|Out-Null;"
        "  $story=$story.NextStoryRange"
        "};"
        "$doc.Save();"
        "$doc.Close();"
        "$word.Quit();"
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            check=True,
            capture_output=True,
            text=True,
            timeout=180,
        )
        return {"attempted": True, "updated": True, "error": None}
    except subprocess.CalledProcessError as e:
        return {
            "attempted": True,
            "updated": False,
            "error": (e.stderr or e.stdout or str(e))[:1200],
        }
    except Exception as e:  # noqa: BLE001
        return {"attempted": True, "updated": False, "error": str(e)}


def convert_docx(
    input_path: Path,
    output_path: Path,
    zotero_db: Path | None,
    connector_base: str,
    add_missing: bool,
    add_bibliography_field: bool,
    replace_reference_list: bool,
    inject_doc_prefs: bool,
    style_id: str,
    citation_pattern: str,
    word_update: bool,
    word_update_visible: bool,
) -> dict:
    with ZipFile(input_path, "r") as zin:
        xml_bytes = zin.read("word/document.xml")
    _register_namespaces_from_xml(xml_bytes)
    original_xml_text = xml_bytes.decode("utf-8", errors="replace")

    root = ET.fromstring(xml_bytes)
    paragraphs = root.findall(f".//{W}p")
    ref_start = _find_references_start(paragraphs)
    if ref_start is None:
        raise RuntimeError("Could not find 'B1. References' heading in document")
    refs = _parse_reference_map(paragraphs, ref_start)
    ref_numbers = set(refs.keys())

    if citation_pattern == "auto":
        selected_citation_pattern, citation_pattern_stats = _detect_citation_pattern(
            paragraphs, ref_start, ref_numbers, include_superscript=True
        )
    elif citation_pattern == "auto-safe":
        selected_citation_pattern, citation_pattern_stats = _detect_citation_pattern(
            paragraphs, ref_start, ref_numbers, include_superscript=False
        )
    else:
        selected_citation_pattern = citation_pattern
        citation_pattern_stats = {}

    index: LocalZoteroIndex | None = None
    matches: dict[int, MatchResult] = {n: MatchResult(item=None, method=None, score=0.0) for n in refs}
    local_match_count = 0
    add_report = {"attempted": 0, "added": 0, "errors": []}

    if zotero_db and zotero_db.exists():
        index = LocalZoteroIndex(zotero_db)
        index.load()
        for n, ref in refs.items():
            m = index.match(ref)
            matches[n] = m
            if m.item:
                local_match_count += 1

        if add_missing:
            missing_refs = [refs[n] for n, m in matches.items() if m.item is None]
            add_report = _add_missing_refs_to_zotero(missing_refs, connector_base)
            if add_report["added"] > 0:
                # Re-load DB and re-match after connector insert.
                index = LocalZoteroIndex(zotero_db)
                index.load()
                local_match_count = 0
                for n, ref in refs.items():
                    m = index.match(ref)
                    matches[n] = m
                    if m.item:
                        local_match_count += 1

    converted_fields = 0
    converted_paras = 0
    bibliography_field_inserted = False
    custom_props_updated = False
    zotero_version = _detect_zotero_version(connector_base) if inject_doc_prefs else None

    for p_idx, para in enumerate(paragraphs):
        if p_idx >= ref_start:
            break
        old_children = list(para)
        new_children: list[ET.Element] = []
        para_changed = False

        for child in old_children:
            if child.tag != f"{W}r":
                new_children.append(child)
                continue
            if child.find(f"{W}instrText") is not None or child.find(f"{W}fldChar") is not None:
                new_children.append(child)
                continue
            t_nodes = child.findall(f"{W}t")
            if len(t_nodes) != 1:
                new_children.append(child)
                continue
            run_text = _get_run_text(child)
            matches_in_run = _find_citation_matches_in_run(child, run_text, selected_citation_pattern)
            if not matches_in_run:
                new_children.append(child)
                continue

            para_changed = True
            last = 0
            for m in matches_in_run:
                start = int(m["start"])
                end = int(m["end"])
                literal = str(m["literal"])
                token = str(m["token"])
                if start > last:
                    pre = run_text[last:start]
                    if pre:
                        new_children.append(_mk_run_with_text(child, pre))
                instruction = _build_citation_instruction(literal, token, refs, matches, index)
                new_children.append(_mk_fldchar_run(child, "begin"))
                new_children.append(_mk_instr_run(child, instruction))
                new_children.append(_mk_fldchar_run(child, "separate"))
                new_children.append(_mk_run_with_text(child, literal))
                new_children.append(_mk_fldchar_run(child, "end"))
                converted_fields += 1
                last = end

            if last < len(run_text):
                tail = run_text[last:]
                if tail:
                    new_children.append(_mk_run_with_text(child, tail))

        if para_changed:
            converted_paras += 1
            for c in old_children:
                para.remove(c)
            for c in new_children:
                para.append(c)

    # Re-scan paragraphs after citation edits before bibliography insertion.
    paragraphs = root.findall(f".//{W}p")
    if add_bibliography_field:
        ref_start = _find_references_start(paragraphs)
        if ref_start is not None:
            bibliography_field_inserted = _insert_or_replace_bibliography_field(
                root=root,
                paragraphs=paragraphs,
                ref_start=ref_start,
                replace_reference_list=replace_reference_list,
            )

    ET.register_namespace("w", W_NS)
    xml_out = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    xml_out_text = xml_out.decode("utf-8", errors="replace")
    xml_out_text = _inject_original_root_namespaces(xml_out_text, original_xml_text)
    xml_out = xml_out_text.encode("utf-8")

    custom_xml_out: bytes | None = None
    effective_style_id = style_id
    if inject_doc_prefs:
        with ZipFile(input_path, "r") as zin:
            existing_custom = zin.read("docProps/custom.xml") if "docProps/custom.xml" in zin.namelist() else None
        if style_id.lower() == "auto":
            existing_style = _extract_style_id_from_custom_props(existing_custom)
            effective_style_id = existing_style or "http://www.zotero.org/styles/nature"
        pref_value = _build_zotero_pref_value(_rand_id(8), effective_style_id, zotero_version or "7.0.0")
        custom_xml_out = _update_custom_props_xml(existing_custom, pref_value)
        custom_props_updated = True

    with ZipFile(input_path, "r") as zin, ZipFile(output_path, "w", compression=ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename == "word/document.xml":
                zout.writestr(info, xml_out)
            elif info.filename == "docProps/custom.xml" and custom_xml_out is not None:
                zout.writestr(info, custom_xml_out)
            else:
                zout.writestr(info, zin.read(info.filename))
        if custom_xml_out is not None and "docProps/custom.xml" not in zin.namelist():
            zout.writestr("docProps/custom.xml", custom_xml_out)

    word_update_report = {"attempted": False, "updated": False, "error": None}
    if word_update:
        word_update_report = _word_update_docx_fields(output_path, visible=word_update_visible)

    matched_numbers = sorted([n for n, m in matches.items() if m.item is not None])
    unmatched_numbers = sorted([n for n, m in matches.items() if m.item is None])
    match_details = {
        str(n): {
            "matched": bool(matches[n].item),
            "method": matches[n].method,
            "score": round(matches[n].score, 4),
            "item_key": matches[n].item.key if matches[n].item else None,
            "item_id": matches[n].item.item_id if matches[n].item else None,
            "item_title": matches[n].item.title if matches[n].item else None,
        }
        for n in sorted(matches.keys())
    }

    return {
        "input": str(input_path),
        "output": str(output_path),
        "reference_count": len(refs),
        "matched_local_zotero": local_match_count,
        "matched_numbers": matched_numbers,
        "unmatched_numbers": unmatched_numbers,
        "connector_add_report": add_report,
        "converted_fields": converted_fields,
        "converted_paragraphs": converted_paras,
        "bibliography_field_inserted": bibliography_field_inserted,
        "custom_props_updated": custom_props_updated,
        "zotero_version_detected": zotero_version,
        "style_id_used": effective_style_id,
        "citation_pattern_selected": selected_citation_pattern,
        "citation_pattern_stats": citation_pattern_stats,
        "word_update_report": word_update_report,
        "local_user_key": index.local_user_key if index else None,
        "zotero_db": str(zotero_db) if zotero_db else None,
        "match_details": match_details,
    }


def _workflow_pattern_order(preferred: str) -> list[str]:
    all_patterns = ["auto-safe", "auto", "paren", "bracket", "superscript"]
    if preferred not in all_patterns:
        return all_patterns
    return [preferred] + [p for p in all_patterns if p != preferred]


def _workflow_quality(report: dict, add_bibliography_field: bool, inject_doc_prefs: bool) -> dict:
    unmatched_count = len(report.get("unmatched_numbers", []))
    converted_fields = int(report.get("converted_fields", 0))
    bibliography_ok = (not add_bibliography_field) or bool(report.get("bibliography_field_inserted"))
    doc_prefs_ok = (not inject_doc_prefs) or bool(report.get("custom_props_updated"))
    fields_ok = converted_fields > 0
    hard_pass = unmatched_count == 0 and bibliography_ok and doc_prefs_ok and fields_ok
    return {
        "hard_pass": hard_pass,
        "unmatched_count": unmatched_count,
        "converted_fields": converted_fields,
        "bibliography_ok": bibliography_ok,
        "doc_prefs_ok": doc_prefs_ok,
        "fields_ok": fields_ok,
    }


def _workflow_is_better(current: dict | None, candidate: dict) -> bool:
    if current is None:
        return True
    candidate_key = (
        int(candidate["hard_pass"]),
        int(candidate["fields_ok"]),
        int(candidate["bibliography_ok"]),
        int(candidate["doc_prefs_ok"]),
        -int(candidate["unmatched_count"]),
        int(candidate["converted_fields"]),
    )
    current_key = (
        int(current["hard_pass"]),
        int(current["fields_ok"]),
        int(current["bibliography_ok"]),
        int(current["doc_prefs_ok"]),
        -int(current["unmatched_count"]),
        int(current["converted_fields"]),
    )
    return candidate_key > current_key


def convert_docx_managed(
    input_path: Path,
    output_path: Path,
    zotero_db: Path | None,
    connector_base: str,
    add_missing: bool,
    add_bibliography_field: bool,
    replace_reference_list: bool,
    inject_doc_prefs: bool,
    style_id: str,
    citation_pattern: str,
    word_update: bool,
    word_update_visible: bool,
) -> dict:
    patterns = _workflow_pattern_order(citation_pattern)
    attempts: list[dict] = []
    best_report: dict | None = None
    best_quality: dict | None = None
    best_output: Path | None = None
    best_attempt_index: int | None = None
    temp_root = Path.cwd() / ".tmp_citationresolver_trials"
    temp_root.mkdir(parents=True, exist_ok=True)
    temp_dir = temp_root / _rand_id(10)
    temp_dir.mkdir(parents=True, exist_ok=True)
    try:
        for idx, pattern in enumerate(patterns, start=1):
            trial_output = temp_dir / f"attempt_{idx}_{pattern}{output_path.suffix}"
            report = convert_docx(
                input_path=input_path,
                output_path=trial_output,
                zotero_db=zotero_db,
                connector_base=connector_base,
                add_missing=add_missing,
                add_bibliography_field=add_bibliography_field,
                replace_reference_list=replace_reference_list,
                inject_doc_prefs=inject_doc_prefs,
                style_id=style_id,
                citation_pattern=pattern,
                word_update=False,
                word_update_visible=False,
            )
            quality = _workflow_quality(report, add_bibliography_field, inject_doc_prefs)
            attempts.append(
                {
                    "attempt": idx,
                    "requested_pattern": pattern,
                    "selected_pattern": report.get("citation_pattern_selected"),
                    "output": str(trial_output),
                    "quality": quality,
                }
            )
            if _workflow_is_better(best_quality, quality):
                best_quality = quality
                best_report = report
                best_output = trial_output
                best_attempt_index = idx
            if quality["hard_pass"]:
                break

        if best_report is None or best_output is None or best_quality is None or best_attempt_index is None:
            raise RuntimeError("Managed workflow failed before producing a candidate output")

        shutil.copy2(best_output, output_path)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    final_report = dict(best_report)
    final_report["output"] = str(output_path)
    final_report["workflow_mode"] = "managed"
    final_report["managed_workflow"] = {
        "enabled": True,
        "preferred_pattern": citation_pattern,
        "attempted_patterns": [a["requested_pattern"] for a in attempts],
        "selected_attempt": best_attempt_index,
        "selected_pattern": final_report.get("citation_pattern_selected"),
        "final_quality": best_quality,
        "attempts": attempts,
    }

    if word_update:
        final_report["word_update_report"] = _word_update_docx_fields(output_path, visible=word_update_visible)

    return final_report


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Managed DOCX citation repair into live Zotero fields")
    p.add_argument(
        "input_path",
        nargs="?",
        help="Input .docx path",
    )
    p.add_argument(
        "--input",
        dest="input_opt",
        help="Input .docx path (backward-compatible flag form)",
    )
    p.add_argument("--output", help="Output .docx path (default: <input_stem>_zotero_integrated.docx)")
    p.add_argument(
        "--zotero-db",
        default=str(Path.home() / "Zotero" / "zotero.sqlite"),
        help="Path to local Zotero sqlite DB (read immutable mode)",
    )
    p.add_argument(
        "--connector-base",
        default="http://127.0.0.1:23119",
        help="Local Zotero connector base URL",
    )
    p.add_argument(
        "--no-add-missing-to-zotero",
        action="store_false",
        dest="add_missing_to_zotero",
        help="Disable adding unmatched bibliography refs to local Zotero via connector/saveItems",
    )
    p.add_argument(
        "--no-add-bibliography-field",
        action="store_false",
        dest="add_bibliography_field",
        help="Disable inserting a live Zotero bibliography field at B1. References",
    )
    p.add_argument(
        "--no-replace-reference-list",
        action="store_false",
        dest="replace_reference_list",
        help="Disable replacing numbered reference list paragraphs when bibliography field is inserted",
    )
    p.add_argument(
        "--no-inject-doc-prefs",
        action="store_false",
        dest="inject_doc_prefs",
        help="Disable injecting Zotero document preferences (enabled by default)",
    )
    p.add_argument(
        "--style-id",
        default="auto",
        help="CSL style id for injected prefs, or 'auto' to reuse existing doc style",
    )
    p.add_argument(
        "--citation-pattern",
        choices=["auto-safe", "auto", "paren", "bracket", "superscript"],
        default="auto-safe",
        help="Preferred first detector; managed workflow may try additional patterns",
    )
    p.add_argument(
        "--word-update",
        action="store_true",
        help="After writing output, open Word via COM and update all fields, then save",
    )
    p.add_argument(
        "--word-update-visible",
        action="store_true",
        help="When --word-update is enabled, keep Word visible while updating",
    )
    p.add_argument(
        "--report-json",
        help="Optional path to write detailed JSON report",
    )
    p.set_defaults(
        inject_doc_prefs=True,
        add_missing_to_zotero=True,
        add_bibliography_field=True,
        replace_reference_list=True,
        word_update=False,
        word_update_visible=False,
    )
    return p


def main() -> int:
    args = build_parser().parse_args()
    input_arg = args.input_opt or args.input_path
    if not input_arg:
        raise ValueError("Provide an input .docx path (positional or --input)")
    input_path = Path(input_arg).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Input not found: {input_path}")

    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        output_path = input_path.with_name(f"{input_path.stem}_zotero_integrated.docx")

    zotero_db = Path(args.zotero_db).expanduser().resolve() if args.zotero_db else None

    report = convert_docx_managed(
        input_path=input_path,
        output_path=output_path,
        zotero_db=zotero_db,
        connector_base=args.connector_base,
        add_missing=args.add_missing_to_zotero,
        add_bibliography_field=args.add_bibliography_field,
        replace_reference_list=args.replace_reference_list,
        inject_doc_prefs=args.inject_doc_prefs,
        style_id=args.style_id,
        citation_pattern=args.citation_pattern,
        word_update=args.word_update,
        word_update_visible=args.word_update_visible,
    )
    if args.report_json:
        Path(args.report_json).expanduser().resolve().write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
