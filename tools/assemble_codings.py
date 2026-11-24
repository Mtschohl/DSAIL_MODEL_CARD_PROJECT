#!/usr/bin/env python3
"""Merge per-chunk JSONL coding files into one validated Excel workbook.

Usage: python tools/assemble_codings.py RUN_DIR CODEBOOK.xlsm OUT.xlsx
RUN_DIR must contain manifest.json, chunks/*.txt and out/*.jsonl.
"""
import datetime as dt
import difflib
import json
import re
import sys
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

run, cb_path, out = Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3])

# ---- codebook ----
wb = openpyxl.load_workbook(cb_path, read_only=True)
rows = [r for r in wb["Codes - Full Set"].iter_rows(values_only=True)]
codebook = pd.DataFrame(
    [
        {
            "code": str(r[2]).strip(),
            "code_type": str(r[0]).strip(),
            "family": str(r[1]).strip(),
            "disclosure_feature": str(r[3]).strip(),
            "description": str(r[4]).strip(),
            "theoretical_framework": str(r[5]).strip() if len(r) > 5 and r[5] else "",
        }
        for r in rows[1:]
        if r and r[2]
    ]
)
CODES = codebook["code"].tolist()
meta = codebook.set_index("code")


def norm(s: str) -> str:
    s = s.replace("’", "'").replace("‘", "'").replace("“", '"').replace("”", '"')
    s = s.replace("–", "-").replace("—", "-").replace("­", "")
    s = re.sub(r"-\s+", "-", s)  # hyphenation across line breaks
    return re.sub(r"\s+", " ", s).strip().lower()


def quote_check(q: str, src_norm: str):
    nq = norm(q)
    if not nq:
        return "empty", 0.0
    if nq in src_norm:
        return "exact", 1.0
    # ellipsis-separated fragments
    frags = [f.strip() for f in re.split(r"\.\.\.|…", nq) if f.strip()]
    if len(frags) > 1 and all(f in src_norm for f in frags):
        return "exact (ellipsis)", 1.0
    sm = difflib.SequenceMatcher(None, src_norm, nq, autojunk=False)
    m = sm.find_longest_match(0, len(src_norm), 0, len(nq))
    lo = max(0, m.a - m.b)
    ratio = difflib.SequenceMatcher(None, src_norm[lo : lo + len(nq) + 30], nq, autojunk=False).ratio()
    return ("near" if ratio >= 0.9 else "not_found"), round(ratio, 3)


manifest = json.loads((run / "manifest.json").read_text())
codings, sections, problems = [], [], []
for m in manifest:
    stem = Path(m["file"]).stem
    src_norm = norm(Path(m["file"]).read_text(encoding="utf-8"))
    jf = run / "out" / f"{stem}.jsonl"
    if not jf.exists():
        problems.append(f"MISSING OUTPUT: {jf.name}")
        continue
    for ln, line in enumerate(jf.read_text(encoding="utf-8").splitlines(), 1):
        if not line.strip():
            continue
        try:
            rec = json.loads(line)
        except json.JSONDecodeError:
            problems.append(f"{jf.name}:{ln} unparseable line")
            continue
        base = {"model_card": m["card"], "chunk": f'part {m["part"]} ({m["pages"]})',
                "page": str(rec.get("page", "")), "section_heading": rec.get("section_heading", "")}
        if rec.get("type") == "section":
            sections.append({**base, "n_codes_reported": rec.get("n_codes", "")})
            continue
        code = str(rec.get("code", "")).strip()
        qc, sim = quote_check(str(rec.get("supporting_text", "")), src_norm)
        codings.append({
            **base,
            "code": code,
            "disclosure_feature": meta.loc[code, "disclosure_feature"] if code in meta.index else "",
            "code_type": meta.loc[code, "code_type"] if code in meta.index else "",
            "family": meta.loc[code, "family"] if code in meta.index else "",
            "supporting_text": rec.get("supporting_text", ""),
            "rationale": rec.get("rationale", ""),
            "code_in_codebook": code in meta.index,
            "quote_check": qc,
            "quote_similarity": sim,
        })

df = pd.DataFrame(codings)
df = df.drop_duplicates(subset=["model_card", "code", "supporting_text"]).reset_index(drop=True)

# ---- dependency rules ----
DEPS = {"PL-B-4.2": "PL-B-4.1", "R-A-19.3": "R-A-19.2", "R-B-22.2": "R-B-22.1"}
keyset = set(zip(df["model_card"], df["code"], df["supporting_text"].map(norm)))
sec_codes = df.groupby(["model_card", "section_heading"])["code"].apply(set).to_dict()


def dep(r):
    parent = DEPS.get(r["code"])
    if not parent:
        return ""
    if (r["model_card"], parent, norm(r["supporting_text"])) in keyset:
        return f"OK ({parent} on same text)"
    if parent in sec_codes.get((r["model_card"], r["section_heading"]), set()):
        return f"OK ({parent} in same section)"
    return f"VIOLATION: {parent} not ticked"


df["dependency_check"] = df.apply(dep, axis=1)
df["needs_review"] = (
    df["quote_check"].eq("not_found") | ~df["code_in_codebook"] | df["dependency_check"].str.startswith("VIOLATION")
)
df.insert(0, "coding_id", range(1, len(df) + 1))

cards = list(dict.fromkeys(m["card"] for m in manifest))
freq = df.groupby(["model_card", "code"]).size().unstack(fill_value=0).reindex(index=cards, columns=CODES, fill_value=0)
freq["TOTAL"] = freq.sum(axis=1)
presence = (freq[CODES] > 0).astype(int)
presence["N_CODES_PRESENT"] = presence.sum(axis=1)
by_code = codebook[["code", "code_type", "family", "disclosure_feature"]].copy()
for c in cards:
    by_code[c] = by_code["code"].map(freq.loc[c]).fillna(0).astype(int)
by_code["TOTAL"] = by_code[cards].sum(axis=1)

sec = pd.DataFrame(sections)
if not sec.empty:
    counted = df.groupby(["model_card", "section_heading"]).size().rename("n_codings_in_workbook")
    sec = sec.merge(counted, on=["model_card", "section_heading"], how="left").fillna({"n_codings_in_workbook": 0})
    sec["n_codings_in_workbook"] = sec["n_codings_in_workbook"].astype(int)

summary = pd.DataFrame({
    "model_card": cards,
    "chunks_coded": [sum(1 for m in manifest if m["card"] == c) for c in cards],
    "sections_recorded": [int((sec["model_card"] == c).sum()) if not sec.empty else 0 for c in cards],
    "total_codings": [int((df["model_card"] == c).sum()) for c in cards],
    "distinct_codes_present": [int(presence.loc[c, "N_CODES_PRESENT"]) for c in cards],
    "quotes_exact": [int(((df["model_card"] == c) & df["quote_check"].str.startswith("exact")).sum()) for c in cards],
    "rows_needing_review": [int(((df["model_card"] == c) & df["needs_review"]).sum()) for c in cards],
})

notes = pd.DataFrame({"item": [
    "Generated", "Codebook", "Task instructions", "Method", "Quote check",
    "Dependency check", "Needs review", "Excluded material", "Processing problems"], "value": [
    dt.datetime.now().isoformat(timespec="minutes"),
    f"{cb_path.name}, sheet 'Codes - Full Set' ({len(CODES)} codes)",
    "Read through each model card provided. Boilerplate and templates have been removed. As you read through each "
    "model card, reference the codebook attached. For each instance, tick the corresponding code, and give the relevant "
    "supporting text. You may tick each code more than once; one instance of supporting text may have more than one "
    "applicable code. Do not skip sections of the model cards.",
    "LLM-assisted coding (Claude, run in a Claude Code session; not on IU Quartz). Each card was split at page/section "
    "boundaries into chunks; each chunk was read in full and coded by a separate model instance using identical instructions.",
    "Every supporting_text was matched against the source chunk after normalizing whitespace, curly quotes and dashes. "
    "exact = verbatim; near = >=90% similar; not_found = could not be located.",
    "PL-B-4.2 requires PL-B-4.1; R-A-19.3 requires R-A-19.2; R-B-22.2 requires R-B-22.1.",
    "TRUE when the quote was not found, the code is not in the codebook, or a dependency rule is violated.",
    "Table-of-contents entries, page numbers and page markers were treated as navigation and not coded. "
    "GPT-5.5 Tables 10, 14 and 16 were not coded because their columns are interleaved in the text extraction and no "
    "verbatim quote could be isolated; these should be coded by hand from the PDF.",
    "; ".join(problems) or "none"]})

out.parent.mkdir(parents=True, exist_ok=True)
with pd.ExcelWriter(out, engine="openpyxl") as xw:
    summary.to_excel(xw, sheet_name="Summary", index=False)
    df.to_excel(xw, sheet_name="Codings", index=False)
    by_code.to_excel(xw, sheet_name="Code_Counts_by_Code", index=False)
    freq.to_excel(xw, sheet_name="Code_Frequency", index_label="model_card")
    presence.to_excel(xw, sheet_name="Code_Presence_0_1", index_label="model_card")
    (sec if not sec.empty else pd.DataFrame()).to_excel(xw, sheet_name="Section_Coverage", index=False)
    codebook.to_excel(xw, sheet_name="Codebook", index=False)
    notes.to_excel(xw, sheet_name="Notes", index=False)
    head = PatternFill("solid", fgColor="DCE6F1")
    for ws in xw.book.worksheets:
        ws.freeze_panes = "B2"
        for c in ws[1]:
            c.font, c.fill = Font(bold=True), head
        for col in ws.columns:
            w = max(len(str(c.value or "")) for c in list(col)[:300])
            ws.column_dimensions[col[0].column_letter].width = min(max(9, w + 2), 60)
        if ws.title in {"Codings", "Notes", "Codebook"}:
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
        if ws.title == "Codings":
            ws.auto_filter.ref = ws.dimensions
print(summary.to_string(index=False))
print("quote_check:", df["quote_check"].value_counts().to_dict())
print("dependency violations:", int(df["dependency_check"].str.startswith("VIOLATION").sum()))
print("problems:", problems or "none")
print("wrote", out)
