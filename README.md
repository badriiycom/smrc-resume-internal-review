# SMRC Resume Internal Review

Tooling to triage the corporate "BD Resumes Active" resume library against
Commence's eight SMRC Key Personnel (KP) role checklists, so recruiting can
find potential SMRC matches among 500+ resumes without a manual read-through
of every file.

See [`smrc_screener/RUNBOOK.md`](smrc_screener/RUNBOOK.md) for full setup and
usage instructions. Short version:

```bash
cd smrc_screener
python -m pip install -r requirements.txt
export ANTHROPIC_API_KEY="sk-ant-..."

# Test run against a synced local copy of the resume folder
python smrc_screen.py --input "/path/to/BD Resumes Active" --limit 5 --out test.xlsx

# Full run — batch mode: ~50% cheaper, submits all resumes as one async job
python smrc_screen.py --input "/path/to/BD Resumes Active" --out SMRC_Resume_Triage_Tracker.xlsx --mode batch --label full-run
```

The script never talks to SharePoint directly — sync or download the "BD
Resumes Active" library to local disk first (see the runbook). It never
uploads resumes anywhere; each resume is sent only to the Anthropic API for
scoring, and results land in a local Excel tracker plus a CSV checkpoint that
lets you stop and resume a long run without re-billing or losing progress.
No admin rights on your machine to install Python? See the runbook's
no-install setup using the Python embeddable package.

Resumes, the checkpoint CSV, and the output `.xlsx` are excluded from git via
`.gitignore` — this repo holds the tool, not candidate data.

## What's in here

| Path | Purpose |
|---|---|
| `smrc_screener/smrc_screen.py` | Main CLI — extracts resume text, scores it against all 8 KP roles via Claude, writes the Excel tracker |
| `smrc_screener/kp_criteria.py` | The eight KP role rulesets (knockouts + strategic differentiators), transcribed from the Commence screening checklists and the SMRC SOW |
| `smrc_screener/requirements.txt` | Python dependencies |
| `smrc_screener/RUNBOOK.md` | Step-by-step setup, running, and troubleshooting guide |

## Updating criteria

If the SOW or a KP checklist changes, edit the relevant role block in
`smrc_screener/kp_criteria.py` — it's plain Python data, one block per role.
