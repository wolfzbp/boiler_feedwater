#!/usr/bin/env python3
"""Pipeline de calidad y estandarizacion para Tela y Color.

Principios:
- Nunca sobreescribe los campos fuente: conserva lineage.
- Separa parsing deterministico de resolucion contra datos maestros.
- Pantone FHI se normaliza por numero base NN-NNNN; para uso textil se propone
  una clave canonica TCX, conservando el sufijo observado (TCX/TPG/TPX).
- Los nombres oficiales Pantone NO se inventan. Pueden cargarse desde un master
  autorizado con --pantone-master.

El script usa artifact_tool para leer/escribir XLSX.
"""
from __future__ import annotations

import argparse
import csv
import re
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from artifact_tool import Blob, SpreadsheetFile

FHI_RE = re.compile(r"(?<!\d)(1[1-9])[-\s]?([0-9]{4})(?!\d)")
FHI_SUFFIX_RE = re.compile(r"\b(TCX|TPG|TPX|TGX)\b")
INTERNAL_6_RE = re.compile(r"(?<!\d)(7\d{5})(?!\d)")
INTERNAL_8_RE = re.compile(r"(?<!\d)(7\d{7})(?!\d)")
PMS_RE = re.compile(r"\bPANTONE\s+([0-9]{2,4})(?:\s*([CU]))?\b")
YEAR_RE = re.compile(r"\b20\d{2}\b")

NON_COLOR_TOKENS = {
    "LAVADO_MAQUINA",
    "LAVADO MAQUINA",
}

FINISH_TERMS = [
    "RAMADO", "CRUDO", "TUBULAR", "ABIERTO", "PERCHADO", "ESMERILADO",
    "COMPACTADO", "LAVADO", "TENIDO", "TEÑIDO", "TERMOFIJADO", "PRETERMOFIJADO",
]


def text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def normalize_text(value: Any) -> str:
    s = text(value).upper().replace("–", "-").replace("—", "-")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s).strip()
    return s


def collapse_repeated_segments(value: Any) -> tuple[str, bool]:
    """Colapsa concatenaciones del mismo valor separadas por coma."""
    s = normalize_text(value)
    parts = [p.strip() for p in s.split(",") if p.strip()]
    if len(parts) <= 1:
        return s, False
    unique = []
    for p in parts:
        if p not in unique:
            unique.append(p)
    return ", ".join(unique), len(unique) < len(parts)


def parse_color(value: Any, pantone_master: dict[str, dict[str, str]] | None = None) -> dict[str, Any]:
    raw = text(value)
    normalized, repeated = collapse_repeated_segments(raw)

    fhi = FHI_RE.search(normalized)
    fhi_base = f"{fhi.group(1)}-{fhi.group(2)}" if fhi else ""
    suffix_m = FHI_SUFFIX_RE.search(normalized)
    source_suffix = suffix_m.group(1) if suffix_m else ""
    int8_m = INTERNAL_8_RE.search(normalized)
    int6_m = INTERNAL_6_RE.search(normalized)
    pms_m = PMS_RE.search(normalized)

    internal_code = int8_m.group(1) if int8_m else (int6_m.group(1) if int6_m else "")
    pms_code = ""
    if pms_m:
        pms_code = pms_m.group(1) + ((" " + pms_m.group(2)) if pms_m.group(2) else "")

    # Extrae un nombre candidato del primer segmento sin destruir el raw.
    name = normalized.split(",")[0].strip()
    name = FHI_RE.sub(" ", name)
    name = FHI_SUFFIX_RE.sub(" ", name)
    name = re.sub(r"\bCOD(?:IGO)?\s*INTERNO\b", " ", name)
    name = INTERNAL_8_RE.sub(" ", name)
    name = INTERNAL_6_RE.sub(" ", name)
    name = PMS_RE.sub(" ", name)
    name = re.sub(r"\bREF\b", " ", name)
    name = YEAR_RE.sub(" ", name)
    name = re.sub(r"\s+", " ", name).strip(" -_/")

    if not raw:
        system = "BLANK"
    elif fhi_base:
        system = "PANTONE_FHI"
    elif pms_code:
        system = "PANTONE_PMS_GRAPHICS"
    elif int8_m:
        system = "INTERNAL_8D"
    elif int6_m:
        system = "INTERNAL_6D"
    else:
        system = "TEXT_OR_OTHER"

    official_name = ""
    official_family = ""
    if fhi_base and pantone_master and fhi_base in pantone_master:
        official_name = pantone_master[fhi_base].get("tcx_name", "")
        official_family = pantone_master[fhi_base].get("color_family", "")

    flags: list[str] = []
    if not raw:
        flags.append("MISSING_COLOR")
    if repeated:
        flags.append("REPEATED_CONCATENATION")
    if normalize_text(raw) in NON_COLOR_TOKENS:
        flags.append("NON_COLOR_VALUE")
    if source_suffix and not fhi_base:
        flags.append("PANTONE_SUFFIX_WITHOUT_FHI_CODE")
    if fhi_base and not source_suffix:
        flags.append("FHI_SUFFIX_MISSING")
    if fhi_base and not official_name:
        flags.append("PANTONE_MASTER_LOOKUP_REQUIRED")
    if system == "TEXT_OR_OTHER" and raw:
        flags.append("UNCONTROLLED_COLOR_TEXT")

    pantone_canonical = f"{fhi_base} TCX" if fhi_base else ""
    if fhi_base:
        canonical_key = f"PANTONE_FHI|{fhi_base}|TCX"
    elif internal_code:
        canonical_key = f"INTERNAL|{internal_code}"
    elif pms_code:
        canonical_key = f"PMS|{pms_code}"
    elif name:
        canonical_key = f"TEXT|{name}"
    else:
        canonical_key = ""

    return {
        "Color_Normalizado": normalized,
        "Color_Sistema_Detectado": system,
        "Pantone_FHI_Base": fhi_base,
        "Pantone_Sufijo_Origen": source_suffix,
        "Pantone_TCX_Canonico": pantone_canonical,
        "Color_Codigo_Interno": internal_code,
        "Pantone_PMS": pms_code,
        "Color_Nombre_Extraido": name,
        "Color_Nombre_Oficial": official_name,
        "Color_Familia_Oficial": official_family,
        "Color_Clave_Canonica": canonical_key,
        "Color_Flags": " | ".join(flags),
    }


def build_fabric_patterns(values: list[Any]) -> tuple[re.Pattern[str], re.Pattern[str], list[str]]:
    # Codigo: prefijo alfabetico + numero; lookahead permite casos como JL-1642RAMADO.
    leading = re.compile(r"^\s*([A-Z]+)\s*[- ]?\s*(\d{3,6}(?:/\d+)?)(?=\D|$)")
    fam_counts: Counter[str] = Counter()
    for value in values:
        m = leading.search(normalize_text(value))
        if m:
            fam_counts[m.group(1)] += 1
    families = [x for x, _ in fam_counts.most_common()]
    if families:
        fam_alt = "|".join(sorted(map(re.escape, families), key=len, reverse=True))
        anywhere = re.compile(rf"\b({fam_alt})\s*[- ]?\s*(\d{{3,6}}(?:/\d+)?)(?=\D|$)")
    else:
        anywhere = leading
    return leading, anywhere, families


def parse_fabric(value: Any, leading: re.Pattern[str], anywhere: re.Pattern[str]) -> dict[str, Any]:
    raw = text(value)
    normalized = normalize_text(raw)
    flags: list[str] = []

    if not raw:
        return {
            "Tela_Normalizada": "", "Tela_Codigo": "", "Tela_Familia_Codigo": "",
            "Tela_Nombre_Candidato": "", "Tela_Acabado": "", "Tela_Ruta": "",
            "Tela_Codigo_Embebido": False, "Tela_Flags": "MISSING_FABRIC"
        }

    m = leading.search(normalized)
    embedded = False
    if not m:
        m = anywhere.search(normalized)
        embedded = bool(m)

    code = ""
    family = ""
    residual = normalized
    if m:
        family = m.group(1)
        code = f"{family}-{m.group(2)}"
        residual = (normalized[:m.start()] + " " + normalized[m.end():]).strip()
        if embedded:
            flags.append("FABRIC_CODE_EMBEDDED")
    else:
        flags.append("FABRIC_CODE_NOT_FOUND")

    finish = ""
    for term in FINISH_TERMS:
        if re.search(rf"\b{re.escape(term)}\b", residual):
            finish = term
            residual = re.sub(rf"\b{re.escape(term)}\b", " ", residual)
            break

    route_m = re.search(r"\bC\s*([0-9]+)\b", residual)
    route = f"C{route_m.group(1)}" if route_m else ""
    residual = re.sub(r"\bC\s*[0-9]+\b", " ", residual)

    # Elimina tokens operativos comunes para dejar un nombre comercial candidato.
    residual = re.sub(r"\bCOTIZ(?:ACION)?\b", " ", residual)
    residual = re.sub(r"\b\d{2,3}%\s*(?:PES|POL|COT|ALG|ELA|ELAS|LYC|SPAN|NYLON)(?:-REC)?\b", " ", residual)
    residual = re.sub(r"\b\d{2,3}[DF]\d+[F]?\b", " ", residual)
    residual = re.sub(r"\s+", " ", residual).strip(" -_/")

    if not finish:
        flags.append("FABRIC_FINISH_MISSING")

    return {
        "Tela_Normalizada": normalized,
        "Tela_Codigo": code,
        "Tela_Familia_Codigo": family,
        "Tela_Nombre_Candidato": residual,
        "Tela_Acabado": finish,
        "Tela_Ruta": route,
        "Tela_Codigo_Embebido": embedded,
        "Tela_Flags": " | ".join(flags),
    }


def load_pantone_master(path: str | None) -> dict[str, dict[str, str]]:
    if not path:
        return {}
    out: dict[str, dict[str, str]] = {}
    with open(path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            key = normalize_text(row.get("pantone_base", ""))
            m = FHI_RE.search(key)
            if not m:
                continue
            base = f"{m.group(1)}-{m.group(2)}"
            out[base] = {
                "tcx_name": text(row.get("tcx_name", "")),
                "color_family": text(row.get("color_family", "")),
            }
    return out


def col_letter(n: int) -> str:
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("input_xlsx")
    ap.add_argument("output_xlsx")
    ap.add_argument("--sheet", default="in")
    ap.add_argument("--pantone-master", default=None,
                    help="CSV autorizado: pantone_base,tcx_name,color_family")
    args = ap.parse_args()

    pantone_master = load_pantone_master(args.pantone_master)
    wb = SpreadsheetFile.import_xlsx(Blob.load(args.input_xlsx))
    ws = wb.worksheets.get_item(args.sheet)

    used = ws.get_range("A1:T1043").values
    headers = used[0]
    records = used[1:]
    index = {h: i for i, h in enumerate(headers)}
    if "Color" not in index or "Tela" not in index:
        raise ValueError("No se encontraron las columnas 'Color' y 'Tela'.")

    fabric_values = [r[index["Tela"]] for r in records]
    leading, anywhere, families = build_fabric_patterns(fabric_values)

    color_fields = [
        "Color_Normalizado", "Color_Sistema_Detectado", "Pantone_FHI_Base",
        "Pantone_Sufijo_Origen", "Pantone_TCX_Canonico", "Color_Codigo_Interno",
        "Pantone_PMS", "Color_Nombre_Extraido", "Color_Nombre_Oficial",
        "Color_Familia_Oficial", "Color_Clave_Canonica", "Color_Flags",
    ]
    fabric_fields = [
        "Tela_Normalizada", "Tela_Codigo", "Tela_Familia_Codigo",
        "Tela_Nombre_Candidato", "Tela_Acabado", "Tela_Ruta",
        "Tela_Codigo_Embebido", "Tela_Flags",
    ]
    new_fields = color_fields + fabric_fields

    enriched_rows: list[list[Any]] = []
    issues: list[list[Any]] = []
    color_raw_counter: Counter[str] = Counter()
    fabric_raw_counter: Counter[str] = Counter()
    color_candidates: dict[str, dict[str, Any]] = {}
    fabric_candidates: dict[str, dict[str, Any]] = {}

    for excel_row, row in enumerate(records, start=2):
        color_raw = row[index["Color"]]
        fabric_raw = row[index["Tela"]]
        c = parse_color(color_raw, pantone_master)
        f = parse_fabric(fabric_raw, leading, anywhere)
        enriched_rows.append([c[k] for k in color_fields] + [f[k] for k in fabric_fields])

        cr = text(color_raw)
        fr = text(fabric_raw)
        color_raw_counter[cr] += 1
        fabric_raw_counter[fr] += 1
        if cr not in color_candidates:
            color_candidates[cr] = c
        if fr not in fabric_candidates:
            fabric_candidates[fr] = f

        for flag in [x.strip() for x in c["Color_Flags"].split("|") if x.strip()]:
            issues.append([excel_row, "Color", cr, flag, c["Color_Clave_Canonica"]])
        for flag in [x.strip() for x in f["Tela_Flags"].split("|") if x.strip()]:
            issues.append([excel_row, "Tela", fr, flag, f["Tela_Codigo"]])

    # Conflictos semanticos para revision de data steward.
    by_fhi = defaultdict(lambda: {"names": Counter(), "suffixes": Counter(), "raw": Counter()})
    by_color_name = defaultdict(lambda: {"fhi": Counter(), "raw": Counter()})
    by_fabric_code = defaultdict(lambda: {"raw": Counter(), "names": Counter(), "finish": Counter(), "route": Counter()})
    for row, enrich in zip(records, enriched_rows):
        c = dict(zip(color_fields, enrich[:len(color_fields)]))
        f = dict(zip(fabric_fields, enrich[len(color_fields):]))
        cr = text(row[index["Color"]])
        fr = text(row[index["Tela"]])
        if c["Pantone_FHI_Base"]:
            d = by_fhi[c["Pantone_FHI_Base"]]
            if c["Color_Nombre_Extraido"]: d["names"][c["Color_Nombre_Extraido"]] += 1
            if c["Pantone_Sufijo_Origen"]: d["suffixes"][c["Pantone_Sufijo_Origen"]] += 1
            d["raw"][cr] += 1
        if c["Color_Nombre_Extraido"]:
            d = by_color_name[c["Color_Nombre_Extraido"]]
            d["fhi"][c["Pantone_FHI_Base"] or "<SIN_FHI>"] += 1
            d["raw"][cr] += 1
        if f["Tela_Codigo"]:
            d = by_fabric_code[f["Tela_Codigo"]]
            d["raw"][fr] += 1
            if f["Tela_Nombre_Candidato"]: d["names"][f["Tela_Nombre_Candidato"]] += 1
            if f["Tela_Acabado"]: d["finish"][f["Tela_Acabado"]] += 1
            if f["Tela_Ruta"]: d["route"][f["Tela_Ruta"]] += 1

    color_conflicts = []
    for code, d in by_fhi.items():
        if len(d["names"]) > 1:
            color_conflicts.append(["SAME_FHI_CODE_MULTIPLE_NAMES", code, " | ".join(f"{k} ({v})" for k,v in d["names"].most_common()), "", " | ".join(f"{k} ({v})" for k,v in d["raw"].most_common(6))])
        if len(d["suffixes"]) > 1:
            color_conflicts.append(["SAME_FHI_CODE_MULTIPLE_SUFFIXES", code, " | ".join(f"{k} ({v})" for k,v in d["suffixes"].most_common()), "", " | ".join(f"{k} ({v})" for k,v in d["raw"].most_common(6))])
    for name, d in by_color_name.items():
        if len(d["fhi"]) > 1:
            color_conflicts.append(["SAME_NAME_MULTIPLE_OR_MISSING_FHI", name, " | ".join(f"{k} ({v})" for k,v in d["fhi"].most_common()), "", " | ".join(f"{k} ({v})" for k,v in d["raw"].most_common(6))])

    fabric_conflicts = []
    for code, d in by_fabric_code.items():
        if len(d["raw"]) > 1:
            fabric_conflicts.append(["SAME_CODE_MULTIPLE_RAW_FORMS", code, " | ".join(f"{k} ({v})" for k,v in d["names"].most_common(6)), " | ".join(f"{k} ({v})" for k,v in d["finish"].most_common()), " | ".join(f"{k} ({v})" for k,v in d["route"].most_common()), " | ".join(f"{k} ({v})" for k,v in d["raw"].most_common(6))])
        if len(d["names"]) > 1:
            fabric_conflicts.append(["SAME_CODE_MULTIPLE_DESCRIPTORS", code, " | ".join(f"{k} ({v})" for k,v in d["names"].most_common(6)), " | ".join(f"{k} ({v})" for k,v in d["finish"].most_common()), " | ".join(f"{k} ({v})" for k,v in d["route"].most_common()), " | ".join(f"{k} ({v})" for k,v in d["raw"].most_common(6))])

    # Anexa columnas sin tocar las columnas fuente.
    start_col = len(headers) + 1
    end_col = start_col + len(new_fields) - 1
    ws.get_range(f"{col_letter(start_col)}1:{col_letter(end_col)}1").values = [new_fields]
    ws.get_range(f"{col_letter(start_col)}2:{col_letter(end_col)}{len(records)+1}").values = enriched_rows
    ws.get_range(f"{col_letter(start_col)}1:{col_letter(end_col)}1").format = {
        "fill": "#243447", "font": {"bold": True, "color": "#FFFFFF"},
        "wrap_text": True, "vertical_alignment": "center"
    }
    ws.get_range(f"{col_letter(start_col)}:{col_letter(end_col)}").format.column_width = 18
    ws.get_range(f"{col_letter(start_col)}:{col_letter(end_col)}").format.wrap_text = True
    ws.freeze_panes.freeze_rows(1)

    # Resumen
    summary = wb.worksheets.get_or_add("DQ_Resumen")
    summary.get_range("A1:B1").values = [["Indicador", "Valor"]]
    summary_rows = [
        ["Registros", len(records)],
        ["Color - vacios", sum(1 for r in records if not text(r[index["Color"]]))],
        ["Color - unicos raw", len(color_raw_counter)],
        ["Color - FHI detectados", sum(1 for row in enriched_rows if row[color_fields.index("Pantone_FHI_Base")])],
        ["Color - concatenaciones repetidas", sum(1 for r in records if collapse_repeated_segments(r[index["Color"]])[1])],
        ["Tela - vacios", sum(1 for r in records if not text(r[index["Tela"]]))],
        ["Tela - unicos raw", len(fabric_raw_counter)],
        ["Tela - codigo detectado", sum(1 for row in enriched_rows if row[len(color_fields)+fabric_fields.index("Tela_Codigo")])],
        ["Tela - familias de codigo detectadas", len(families)],
        ["Issues generados", len(issues)],
        ["Color - conflictos semanticos", len(color_conflicts)],
        ["Tela - conflictos semanticos", len(fabric_conflicts)],
        ["Pantone master cargado", "SI" if pantone_master else "NO"],
    ]
    summary.get_range(f"A2:B{len(summary_rows)+1}").values = summary_rows
    summary.get_range("A1:B1").format = {"fill":"#243447","font":{"bold":True,"color":"#FFFFFF"}}
    summary.get_range("A:B").format.column_width = 32

    # Master color candidato
    mc = wb.worksheets.get_or_add("DQ_Master_Color")
    mc_headers = ["Color_Raw", "Frecuencia"] + color_fields
    mc_rows = [[raw, color_raw_counter[raw]] + [obj[k] for k in color_fields]
               for raw, obj in sorted(color_candidates.items(), key=lambda kv: (-color_raw_counter[kv[0]], kv[0]))]
    mc.get_range(f"A1:{col_letter(len(mc_headers))}1").values = [mc_headers]
    if mc_rows:
        mc.get_range(f"A2:{col_letter(len(mc_headers))}{len(mc_rows)+1}").values = mc_rows
    mc.get_range(f"A1:{col_letter(len(mc_headers))}1").format = {"fill":"#243447","font":{"bold":True,"color":"#FFFFFF"},"wrap_text":True}
    mc.freeze_panes.freeze_rows(1)
    mc.get_range(f"A:{col_letter(len(mc_headers))}").format.column_width = 20
    mc.get_range("A:A").format.column_width = 42
    mc.get_range(f"A:{col_letter(len(mc_headers))}").format.wrap_text = True

    # Master tela candidato
    mt = wb.worksheets.get_or_add("DQ_Master_Tela")
    mt_headers = ["Tela_Raw", "Frecuencia"] + fabric_fields
    mt_rows = [[raw, fabric_raw_counter[raw]] + [obj[k] for k in fabric_fields]
               for raw, obj in sorted(fabric_candidates.items(), key=lambda kv: (-fabric_raw_counter[kv[0]], kv[0]))]
    mt.get_range(f"A1:{col_letter(len(mt_headers))}1").values = [mt_headers]
    if mt_rows:
        mt.get_range(f"A2:{col_letter(len(mt_headers))}{len(mt_rows)+1}").values = mt_rows
    mt.get_range(f"A1:{col_letter(len(mt_headers))}1").format = {"fill":"#243447","font":{"bold":True,"color":"#FFFFFF"},"wrap_text":True}
    mt.freeze_panes.freeze_rows(1)
    mt.get_range(f"A:{col_letter(len(mt_headers))}").format.column_width = 20
    mt.get_range("A:A").format.column_width = 48
    mt.get_range(f"A:{col_letter(len(mt_headers))}").format.wrap_text = True

    # Issues para data steward
    iss = wb.worksheets.get_or_add("DQ_Issues")
    issue_headers = ["Fila_Excel", "Campo", "Valor_Raw", "Issue", "Clave_Candidata"]
    iss.get_range("A1:E1").values = [issue_headers]
    if issues:
        iss.get_range(f"A2:E{len(issues)+1}").values = issues
    iss.get_range("A1:E1").format = {"fill":"#9B2C2C","font":{"bold":True,"color":"#FFFFFF"}}
    iss.freeze_panes.freeze_rows(1)
    iss.get_range("A:E").format.column_width = 24
    iss.get_range("C:C").format.column_width = 48
    iss.get_range("A:E").format.wrap_text = True

    # Conflictos agregados (mucho mas utiles que revisar fila por fila).
    cc = wb.worksheets.get_or_add("DQ_Conflictos_Color")
    cc_headers = ["Tipo_Conflicto", "Clave", "Detalle", "Resolucion", "Ejemplos_Raw"]
    cc.get_range("A1:E1").values = [cc_headers]
    if color_conflicts:
        cc.get_range(f"A2:E{len(color_conflicts)+1}").values = color_conflicts
    cc.get_range("A1:E1").format = {"fill":"#6B46C1","font":{"bold":True,"color":"#FFFFFF"},"wrap_text":True}
    cc.freeze_panes.freeze_rows(1)
    cc.get_range("A:E").format.column_width = 28
    cc.get_range("C:E").format.column_width = 48
    cc.get_range("A:E").format.wrap_text = True

    ct = wb.worksheets.get_or_add("DQ_Conflictos_Tela")
    ct_headers = ["Tipo_Conflicto", "Tela_Codigo", "Descriptores", "Acabados", "Rutas", "Ejemplos_Raw"]
    ct.get_range("A1:F1").values = [ct_headers]
    if fabric_conflicts:
        ct.get_range(f"A2:F{len(fabric_conflicts)+1}").values = fabric_conflicts
    ct.get_range("A1:F1").format = {"fill":"#2B6CB0","font":{"bold":True,"color":"#FFFFFF"},"wrap_text":True}
    ct.freeze_panes.freeze_rows(1)
    ct.get_range("A:F").format.column_width = 26
    ct.get_range("C:F").format.column_width = 46
    ct.get_range("A:F").format.wrap_text = True

    SpreadsheetFile.export_xlsx(wb).save(args.output_xlsx)
    print(f"OK: {args.output_xlsx}")


if __name__ == "__main__":
    main()
