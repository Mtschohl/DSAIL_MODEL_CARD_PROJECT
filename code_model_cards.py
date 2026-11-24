#!/usr/bin/env python3
"""
Qualitative coding of AI model cards against a codebook, using a locally hosted
open-weights LLM (designed for IU Quartz GPU nodes).

Task given to the model (verbatim from the project lead):
  Read through each model card provided. Boilerplate and templates have been
  removed. As you read through each model card, reference the codebook attached.
  For each instance, tick the corresponding code, and give the relevant
  supporting text. You may tick each code more than once; one instance of
  supporting text may have more than one applicable code. Do not skip sections
  of the model cards. When you finish, export the data as an excel spreadsheet.

Coverage guarantee: every model card is split into its sections, and every
section is sent to the model individually. The Section_Coverage sheet records
each section so that nothing is skipped silently.

Usage:
  python code_model_cards.py --cards data/model_cards --codebook data/codebook/codebook.docx \
      --out output/model_card_coding.xlsx --backend vllm --model Qwen/Qwen2.5-32B-Instruct
"""
from __future__ import annotations

import argparse
import datetime as dt
import difflib
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

TASK_INSTRUCTIONS = (
    "Read through each model card provided. Boilerplate and templates have been removed. "
    "As you read through each model card, reference the codebook attached. For each instance, "
    "tick the corresponding code, and give the relevant supporting text. You may tick each code "
    "more than once; one instance of supporting text may have more than one applicable code. "
    "Do not skip sections of the model cards."
)

CARD_EXTS = {".md", ".markdown", ".txt", ".docx", ".pdf", ".html", ".htm"}


# --------------------------------------------------------------------------- #
# Reading files
# --------------------------------------------------------------------------- #
def read_text(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".docx":
        import docx

        d = docx.Document(str(path))
        lines = []
        for p in d.paragraphs:
            style = (p.style.name or "").lower() if p.style is not None else ""
            m = re.match(r"heading (\d)", style)
            if m and p.text.strip():
                lines.append("#" * int(m.group(1)) + " " + p.text.strip())
            elif style == "title" and p.text.strip():
                lines.append("# " + p.text.strip())
            else:
                lines.append(p.text)
        for t in d.tables:
            for row in t.rows:
                lines.append(" | ".join(c.text.strip() for c in row.cells))
        return "\n".join(lines)
    if ext == ".pdf":
        import pdfplumber

        with pdfplumber.open(str(path)) as pdf:
            return "\n".join((pg.extract_text() or "") for pg in pdf.pages)
    if ext in {".html", ".htm"}:
        raw = path.read_text(encoding="utf-8", errors="replace")
        raw = re.sub(r"<h([1-6])[^>]*>", lambda m: "\n" + "#" * int(m.group(1)) + " ", raw, flags=re.I)
        raw = re.sub(r"<(script|style).*?</\1>", "", raw, flags=re.S | re.I)
        raw = re.sub(r"<br\s*/?>|</p>|</li>|</h[1-6]>", "\n", raw, flags=re.I)
        return re.sub(r"<[^>]+>", "", raw)
    return path.read_text(encoding="utf-8", errors="replace")


@dataclass
class Codebook:
    text: str  # full codebook text given to the model
    codes: list[dict] = field(default_factory=list)  # [{"code":..., "definition":...}] if tabular

    @property
    def code_names(self) -> list[str]:
        return [c["code"] for c in self.codes]


def load_codebook(path: Path) -> Codebook:
    ext = path.suffix.lower()
    if ext in {".xlsx", ".xls", ".csv", ".tsv"}:
        if ext == ".csv":
            df = pd.read_csv(path)
        elif ext == ".tsv":
            df = pd.read_csv(path, sep="\t")
        else:
            df = pd.read_excel(path)
        df = df.dropna(how="all")
        cols = {c.lower().strip(): c for c in df.columns}
        code_col = next((cols[k] for k in cols if k in {"code", "code name", "code_name", "label", "name"}), df.columns[0])
        codes = []
        for _, r in df.iterrows():
            name = str(r[code_col]).strip()
            if not name or name.lower() == "nan":
                continue
            rest = "; ".join(f"{c}: {r[c]}" for c in df.columns if c != code_col and pd.notna(r[c]))
            codes.append({"code": name, "definition": rest})
        text = "\n".join(f"- {c['code']}: {c['definition']}" for c in codes)
        return Codebook(text=text, codes=codes)
    # Narrative codebook (docx / pdf / md / txt). Also try to pull code names from docx tables.
    text = read_text(path)
    codes: list[dict] = []
    if ext == ".docx":
        import docx

        for t in docx.Document(str(path)).tables:
            rows = [[c.text.strip() for c in row.cells] for row in t.rows]
            if len(rows) < 2:
                continue
            header = [h.lower() for h in rows[0]]
            idx = next((i for i, h in enumerate(header) if h in {"code", "code name", "label", "name"}), None)
            if idx is None:
                continue
            for row in rows[1:]:
                if row[idx]:
                    rest = "; ".join(f"{rows[0][j]}: {v}" for j, v in enumerate(row) if j != idx and v)
                    codes.append({"code": row[idx], "definition": rest})
    return Codebook(text=text, codes=codes)


# --------------------------------------------------------------------------- #
# Sectioning
# --------------------------------------------------------------------------- #
@dataclass
class Section:
    card: str
    index: int
    heading: str
    text: str


HEADING_RE = re.compile(r"^\s{0,3}(#{1,6})\s+(.*?)\s*#*\s*$")


def split_sections(card_name: str, text: str, max_chars: int) -> list[Section]:
    text = text.replace("\r\n", "\n")
    blocks: list[tuple[str, list[str]]] = [("(Preamble / untitled)", [])]
    for line in text.split("\n"):
        m = HEADING_RE.match(line)
        if m:
            blocks.append((m.group(2).strip() or "(untitled)", []))
        else:
            blocks[-1][1].append(line)

    # Fall back to paragraph chunks when the card has no headings.
    if len(blocks) == 1:
        paras = [p for p in re.split(r"\n\s*\n", text) if p.strip()]
        blocks = [("Part", [p]) for p in paras]

    sections: list[Section] = []
    for heading, lines in blocks:
        body = "\n".join(lines).strip()
        if not body:
            continue
        for k, chunk in enumerate(chunk_text(body, max_chars)):
            h = heading if k == 0 else f"{heading} (cont. {k + 1})"
            sections.append(Section(card_name, len(sections) + 1, h, chunk))

    # Merge consecutive tiny untitled paragraph parts so the model sees reasonable context.
    if sections and all(s.heading.startswith("Part") for s in sections):
        merged, buf = [], ""
        for s in sections:
            if len(buf) + len(s.text) > max_chars and buf:
                merged.append(buf)
                buf = ""
            buf = (buf + "\n\n" + s.text).strip()
        if buf:
            merged.append(buf)
        sections = [Section(card_name, i + 1, f"Part {i + 1}", t) for i, t in enumerate(merged)]
    return sections


def chunk_text(body: str, max_chars: int) -> list[str]:
    if len(body) <= max_chars:
        return [body]
    paras = re.split(r"(\n\s*\n)", body)
    chunks, buf = [], ""
    for p in paras:
        if len(buf) + len(p) > max_chars and buf.strip():
            chunks.append(buf.strip())
            buf = ""
        while len(p) > max_chars:  # a single giant paragraph
            chunks.append(p[:max_chars].strip())
            p = p[max_chars:]
        buf += p
    if buf.strip():
        chunks.append(buf.strip())
    return chunks


# --------------------------------------------------------------------------- #
# Prompting
# --------------------------------------------------------------------------- #
SYSTEM_PROMPT = (
    "You are a careful qualitative research assistant performing deductive content analysis. "
    "You apply a codebook to text exactly as the codebook defines each code. You quote supporting "
    "text verbatim from the source and never paraphrase or invent quotations."
)


def build_prompt(cb: Codebook, sec: Section) -> str:
    code_list = ""
    if cb.codes:
        code_list = (
            "\nAllowed code labels (use these exact labels):\n"
            + "\n".join(f"- {c}" for c in cb.code_names)
            + "\n"
        )
    return f"""TASK
{TASK_INSTRUCTIONS}

CODEBOOK
<<<CODEBOOK
{cb.text}
CODEBOOK>>>
{code_list}
You are now reading model card "{sec.card}", section {sec.index}: "{sec.heading}".
Read the entire section. Identify every instance where a code from the codebook applies.

Rules:
1. "supporting_text" must be copied verbatim from the section below (an exact substring), long enough to justify the code, usually one sentence or clause.
2. A code may be applied multiple times in the section, once per distinct instance.
3. The same supporting_text may receive several codes; output one entry per code.
4. Use the code label exactly as written in the codebook.
5. If no code applies anywhere in the section, return an empty list.

SECTION TEXT
<<<SECTION
{sec.text}
SECTION>>>

Respond with JSON only, in this format:
{{"codings": [{{"code": "<code label>", "supporting_text": "<verbatim quote>", "rationale": "<one short sentence>"}}]}}"""


def parse_json(raw: str) -> list[dict] | None:
    raw = raw.strip()
    raw = re.sub(r"^```(?:json)?|```$", "", raw, flags=re.M).strip()
    start, end = raw.find("{"), raw.rfind("}")
    if start == -1 or end == -1:
        return None
    try:
        obj = json.loads(raw[start : end + 1])
    except json.JSONDecodeError:
        return None
    items = obj.get("codings") if isinstance(obj, dict) else None
    if not isinstance(items, list):
        return None
    return [i for i in items if isinstance(i, dict) and i.get("code")]


# --------------------------------------------------------------------------- #
# Backends
# --------------------------------------------------------------------------- #
class Backend:
    def generate(self, prompts: list[str]) -> list[str]:
        raise NotImplementedError


class VLLMBackend(Backend):
    def __init__(self, model: str, max_new_tokens: int, max_model_len: int, tp: int, dtype: str):
        from vllm import LLM, SamplingParams

        self.llm = LLM(
            model=model,
            tensor_parallel_size=tp,
            dtype=dtype,
            max_model_len=max_model_len,
            enable_prefix_caching=True,  # the codebook prefix is shared by every prompt
            gpu_memory_utilization=0.90,
        )
        self.params = SamplingParams(temperature=0.0, max_tokens=max_new_tokens)

    def generate(self, prompts):
        msgs = [[{"role": "system", "content": SYSTEM_PROMPT}, {"role": "user", "content": p}] for p in prompts]
        outs = self.llm.chat(msgs, self.params, use_tqdm=True)
        return [o.outputs[0].text for o in outs]


class HFBackend(Backend):
    def __init__(self, model: str, max_new_tokens: int, dtype: str):
        import torch
        from transformers import AutoModelForCausalLM, AutoTokenizer

        self.tok = AutoTokenizer.from_pretrained(model)
        torch_dtype = {"float16": torch.float16, "bfloat16": torch.bfloat16, "float32": torch.float32}.get(dtype, "auto")
        self.model = AutoModelForCausalLM.from_pretrained(model, torch_dtype=torch_dtype, device_map="auto")
        self.max_new_tokens = max_new_tokens

    def generate(self, prompts):
        import torch

        outs = []
        for p in prompts:
            msgs = [{"role": "system", "content": SYSTEM_PROMPT}, {"role": "user", "content": p}]
            ids = self.tok.apply_chat_template(msgs, add_generation_prompt=True, return_tensors="pt").to(self.model.device)
            mask = torch.ones_like(ids)
            with torch.no_grad():
                gen = self.model.generate(ids, attention_mask=mask, max_new_tokens=self.max_new_tokens,
                                          do_sample=False, pad_token_id=self.tok.eos_token_id)
            outs.append(self.tok.decode(gen[0, ids.shape[1] :], skip_special_tokens=True))
        return outs


# --------------------------------------------------------------------------- #
# Validation helpers
# --------------------------------------------------------------------------- #
def norm(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip().lower()


def quote_match(quote: str, source: str) -> tuple[str, float]:
    q, s = norm(quote), norm(source)
    if not q:
        return "empty", 0.0
    if q in s:
        return "exact", 1.0
    sm = difflib.SequenceMatcher(None, s, q, autojunk=False)
    m = sm.find_longest_match(0, len(s), 0, len(q))
    lo = max(0, m.a - m.b)
    window = s[lo : lo + len(q) + 20]
    ratio = difflib.SequenceMatcher(None, window, q, autojunk=False).ratio()
    return ("near" if ratio >= 0.85 else "not_found"), round(ratio, 3)


def resolve_code(code: str, cb: Codebook) -> tuple[str, bool]:
    if not cb.codes:
        return code.strip(), True
    lookup = {norm(c): c for c in cb.code_names}
    if norm(code) in lookup:
        return lookup[norm(code)], True
    close = difflib.get_close_matches(norm(code), list(lookup), n=1, cutoff=0.85)
    if close:
        return lookup[close[0]], True
    return code.strip(), False


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--cards", required=True, type=Path, help="Folder of model cards (md/txt/docx/pdf/html)")
    ap.add_argument("--codebook", required=True, type=Path, help="Codebook file (docx/pdf/xlsx/csv/md/txt)")
    ap.add_argument("--out", default=Path("output/model_card_coding.xlsx"), type=Path)
    ap.add_argument("--backend", choices=["vllm", "hf"], default="vllm")
    ap.add_argument("--model", default="Qwen/Qwen2.5-32B-Instruct")
    ap.add_argument("--tp", type=int, default=1, help="Tensor parallel size (number of GPUs) for vLLM")
    ap.add_argument("--dtype", default="float16", help="float16 for V100 GPUs, bfloat16 for A100/H100")
    ap.add_argument("--max-section-chars", type=int, default=6000)
    ap.add_argument("--max-new-tokens", type=int, default=2048)
    ap.add_argument("--max-model-len", type=int, default=16384)
    ap.add_argument("--retries", type=int, default=2, help="Re-ask sections whose output was not valid JSON")
    ap.add_argument("--limit", type=int, default=0, help="Only process the first N cards (for testing)")
    args = ap.parse_args(argv)

    cb = load_codebook(args.codebook)
    print(f"Codebook loaded: {len(cb.text):,} chars; {len(cb.codes)} codes parsed from tables", flush=True)

    card_files = sorted(p for p in args.cards.rglob("*") if p.suffix.lower() in CARD_EXTS and p.is_file())
    if args.limit:
        card_files = card_files[: args.limit]
    if not card_files:
        sys.exit(f"No model cards found in {args.cards}")

    sections: list[Section] = []
    for f in card_files:
        sections.extend(split_sections(f.stem, read_text(f), args.max_section_chars))
    print(f"{len(card_files)} model cards -> {len(sections)} sections", flush=True)

    if args.backend == "vllm":
        backend = VLLMBackend(args.model, args.max_new_tokens, args.max_model_len, args.tp, args.dtype)
    else:
        backend = HFBackend(args.model, args.max_new_tokens, args.dtype)

    prompts = [build_prompt(cb, s) for s in sections]
    raws = backend.generate(prompts)
    results: list[list[dict] | None] = [parse_json(r) for r in raws]

    for attempt in range(args.retries):
        bad = [i for i, r in enumerate(results) if r is None]
        if not bad:
            break
        print(f"Retry {attempt + 1}: {len(bad)} sections returned unparseable output", flush=True)
        redo = backend.generate([prompts[i] + "\n\nReturn ONLY valid JSON." for i in bad])
        for i, r in zip(bad, redo):
            raws[i] = r
            results[i] = parse_json(r)

    coding_rows, coverage_rows = [], []
    for sec, res, raw in zip(sections, results, raws):
        status = "coded" if res else ("no codes applied" if res == [] else "FAILED - check raw output")
        coverage_rows.append(
            {
                "model_card": sec.card,
                "section_no": sec.index,
                "section_heading": sec.heading,
                "section_chars": len(sec.text),
                "n_codes_applied": len(res or []),
                "status": status,
                "section_text": sec.text,
                "raw_model_output": raw if res is None else "",
            }
        )
        for item in res or []:
            code, in_cb = resolve_code(str(item.get("code", "")), cb)
            quote = str(item.get("supporting_text", ""))
            match, score = quote_match(quote, sec.text)
            coding_rows.append(
                {
                    "model_card": sec.card,
                    "section_no": sec.index,
                    "section_heading": sec.heading,
                    "code": code,
                    "supporting_text": quote,
                    "rationale": item.get("rationale", ""),
                    "code_in_codebook": in_cb,
                    "quote_check": match,
                    "quote_similarity": score,
                    "needs_review": (match == "not_found") or (not in_cb),
                }
            )

    codings = pd.DataFrame(
        coding_rows,
        columns=["model_card", "section_no", "section_heading", "code", "supporting_text", "rationale",
                 "code_in_codebook", "quote_check", "quote_similarity", "needs_review"],
    )
    coverage = pd.DataFrame(coverage_rows)

    # Card x code frequency matrix ("ticks"), including codes never applied.
    all_codes = list(dict.fromkeys(cb.code_names + codings["code"].tolist()))
    freq = (
        codings.groupby(["model_card", "code"]).size().unstack(fill_value=0)
        if not codings.empty else pd.DataFrame()
    )
    freq = freq.reindex(index=[f.stem for f in card_files], columns=all_codes, fill_value=0)
    freq.index.name = "model_card"
    presence = (freq > 0).astype(int)

    run_info = pd.DataFrame(
        {
            "item": ["run_timestamp", "model", "backend", "dtype", "n_model_cards", "n_sections",
                     "n_codings", "n_sections_failed", "n_codings_needing_review", "codebook_file", "task_instructions"],
            "value": [dt.datetime.now().isoformat(timespec="seconds"), args.model, args.backend, args.dtype,
                      len(card_files), len(sections), len(codings),
                      int((coverage["status"].str.startswith("FAILED")).sum()),
                      int(codings["needs_review"].sum()) if not codings.empty else 0,
                      str(args.codebook), TASK_INSTRUCTIONS],
        }
    )

    args.out.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(args.out, engine="openpyxl") as xw:
        codings.to_excel(xw, sheet_name="Codings", index=False)
        freq.to_excel(xw, sheet_name="Code_Frequency")
        presence.to_excel(xw, sheet_name="Code_Presence_0_1")
        coverage.to_excel(xw, sheet_name="Section_Coverage", index=False)
        pd.DataFrame(cb.codes or [{"code": "(narrative codebook)", "definition": cb.text}]).to_excel(
            xw, sheet_name="Codebook", index=False)
        run_info.to_excel(xw, sheet_name="Run_Info", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "B2" if ws.title.startswith("Code_") else "A2"
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 70)
    print(f"Wrote {args.out}  ({len(codings)} codings across {len(sections)} sections)", flush=True)


if __name__ == "__main__":
    main()
