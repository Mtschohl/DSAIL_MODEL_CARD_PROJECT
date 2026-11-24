# DSAIL Model Card Project: Automated Codebook Coding

This pipeline applies the project codebook to a corpus of AI model cards with an open-weights LLM
(Qwen2.5-32B-Instruct) on IU Quartz GPU nodes. It exports the results as an Excel workbook.

## Task given to the model
> Read through each model card provided. Boilerplate and templates have been removed. As you read through
> each model card, reference the codebook attached. For each instance, tick the corresponding code, and give
> the relevant supporting text. You may tick each code more than once; one instance of supporting text may
> have more than one applicable code. Do not skip sections of the model cards.

To make sure no section is skipped, the script splits every card at its headings and sends each section to the
model separately. The `Section_Coverage` sheet lists every section and how many codes it received.

## Inputs
- `data/codebook/`: one codebook file (`.docx`, `.pdf`, `.xlsx`, `.csv`, `.md` or `.txt`). A table with a
  `Code` column lets the script check every code label against the codebook.
- `data/model_cards/`: the model cards (`.md`, `.txt`, `.docx`, `.pdf` or `.html`), one file per card.
  Markdown headings (`#`) or Word heading styles mark the sections.

## Running on Quartz
```bash
ssh <username>@quartz.uits.iu.edu
git clone git@github.com:Mtschohl/DSAIL_MODEL_CARD_PROJECT.git
cd DSAIL_MODEL_CARD_PROJECT
bash quartz/setup_env.sh          # one time: environment + model download (~65 GB to /N/slate)
```
Edit `--account` in `quartz/run_coding.slurm`, then:
```bash
sbatch quartz/run_coding.slurm
squeue -u $USER                    # check progress; log is output/slurm-<jobid>.out
```
For a quick trial, add `--limit 3` to the python command in the Slurm file.

## Output workbook (`output/model_card_coding_<date>.xlsx`)
| Sheet | Contents |
|---|---|
| Codings | One row per code application: card, section, code, verbatim supporting text, rationale, review flags |
| Code_Frequency | Card × code counts of how many times each code was ticked |
| Code_Presence_0_1 | Card × code, 1 if the code appears at least once |
| Section_Coverage | Every section of every card, number of codes applied, status |
| Codebook | The codebook as parsed |
| Run_Info | Model, date, counts, and the task instructions |

`quote_check` reports whether the supporting text appears verbatim in the section (`exact`, `near`,
`not_found`). Rows with `needs_review = TRUE` have a quote that was not found or a code label outside the codebook.
LLM coding should be validated against a human-coded subsample (for example, Cohen's kappa) before reporting.
