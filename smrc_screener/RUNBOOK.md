# SMRC KP Resume Screener — Runbook

**Audience:** Badri / IT (whoever runs the script). No prior familiarity with the
SMRC capture needed. Estimated setup time: ~30 min. Full 580-resume run: roughly
1–2 hours unattended, depending on rate limits.

**What it does:** Reads every resume in a folder, sends each to Claude via the
Anthropic API, scores it against all eight SMRC Key Personnel checklists in one
pass, and writes a branded Excel tracker (candidate name, primary KP category,
whether they meet minimum SOW requirements, why, multi-role flags, OCI flags).

---

## 1. Prerequisites

- **Python 3.9+** installed (`python --version`).
- **An Anthropic API key.** Get one from the Anthropic Console
  (https://console.anthropic.com → API Keys). This is a paid API; see cost note
  at the bottom. Do NOT paste the key into the script or SharePoint — set it as
  an environment variable (step 3).
- **The resume files on local disk.** The script reads a *folder*, not SharePoint
  directly. Sync the SharePoint "BD Resumes Active" library to a local folder
  first — either via the OneDrive/SharePoint sync client (right-click the
  library → "Sync"/"Add shortcut to OneDrive") or by selecting all and
  downloading. Point the script at that local folder.
- **(Optional) LibreOffice** on PATH — only needed if you have legacy `.doc` or
  `.rtf` resumes. `.pdf`, `.docx`, and `.txt` work without it.

## 2. Install

```bash
cd smrc_screener
python -m pip install -r requirements.txt
```

## 3. Set the API key (never hard-code it)

macOS / Linux:
```bash
export ANTHROPIC_API_KEY="sk-ant-..."
```
Windows (PowerShell):
```powershell
$env:ANTHROPIC_API_KEY="sk-ant-..."
```

## 4. Do a small test run first (5 resumes)

Confirms extraction, API access, and output format before spending on all 580:
```bash
python smrc_screen.py --input "C:\path\to\resumes" --limit 5 --out test.xlsx
```
Open `test.xlsx`, sanity-check a couple of rows against the actual resumes. If
categories look right, proceed to the full run.

## 5. Full run

```bash
python smrc_screen.py --input "C:\path\to\resumes" --out SMRC_Resume_Triage_Tracker.xlsx --batch full-run
```

- Progress prints per file: `[42/580] jsmith_resume.pdf`.
- Every result is also appended to `screen_checkpoint.csv` as it goes.
- **Interrupted?** Just re-run the exact same command. It reads the checkpoint
  and skips everything already done — no duplicate API charges, no lost work.
- To force a clean re-run, delete `screen_checkpoint.csv` first.

## 6. Output

- **`SMRC_Resume_Triage_Tracker.xlsx`**
  - *Candidate Tracker* tab — one row per resume: name, title/employer, primary
    KP category, meets-minimums (Y/N), fit strength, cited rationale, multi-role
    flag, OCI flag, date.
  - *Category Summary* tab — live counts per role (auto-calculates in Excel).
- Hand this file back to Rachel. It matches the tracker structure already in use.

## 7. Handling candidate data

Resumes and the tracker/checkpoint files contain personal information (names,
contact details, license numbers). Keep them on local disk or the approved
SharePoint libraries only:

- Do **not** commit resumes, `screen_checkpoint.csv`, or the output `.xlsx`
  into this git repository. The `.gitignore` at the repo root already excludes
  the common output filenames and any `resumes/` folder, but double-check
  `git status` before committing if you rename anything.
- Delete local resume copies and checkpoints when the screening pass is done,
  unless your team's retention policy says otherwise.

---

## How the scoring works (so you can explain it / trust it)

- The eight roles' **knockout criteria** (from the Commence checklists + June 3,
  2026 SOW) live in `kp_criteria.py`. Each resume must clear *all* knockouts for
  a role to qualify for it.
- The model is instructed to be **conservative**: it only credits a requirement
  when the resume shows specific evidence (a license, a degree, stated years, a
  named program). Missing/ambiguous = not met. It does not infer.
- **Sanction/exclusion checks** (OIG LEIE + SAM.gov) can't be done from a resume.
  The model flags "sanction/exclusion check pending" in the rationale rather than
  passing or failing on it — your team runs that verification separately.
- **OCI:** resumes showing current employment at Noridian (SMRC incumbent) or
  other named competitors get an OCI flag for Bryan/Albie to review.
- This is a **first-pass triage**, not a hiring decision. Treat "Strong" primary
  matches — especially for the critical-path roles (CIO, SSO, MRM Task 2) — as a
  shortlist for human review, not a final call.

## Updating the criteria

If the SOW or a checklist changes, edit the relevant role block in
`kp_criteria.py` (plain text, clearly labeled per role) and re-run. Nothing else
needs to change.

## Cost & rate limits (rough)

- Model: `claude-sonnet-4-6` (set at the top of `smrc_screen.py`, override with
  `--model` if needed).
- ~580 resumes × one call each. Cost depends on resume length; the script caps
  each resume at ~24k characters to keep it bounded. Ballpark: low tens of
  dollars total, not hundreds. Check current pricing in the Anthropic Console.
- `--sleep 0.5` adds a half-second between calls as a rate-limit cushion. If you
  hit rate-limit errors, raise it (e.g. `--sleep 1.5`). Errors on individual
  files are logged in the `error` column of the checkpoint/tracker and can be
  re-run (delete that row from `screen_checkpoint.csv` and re-run the same
  command).

## Files in this package

| File | Purpose |
|---|---|
| `smrc_screen.py` | Main script — run this |
| `kp_criteria.py` | The eight KP rulesets (edit if SOW changes) |
| `requirements.txt` | Python dependencies |
| `RUNBOOK.md` | This file |
