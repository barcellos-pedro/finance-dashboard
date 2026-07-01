# Agent Instructions for this repository

Purpose: Help AI coding agents be immediately productive when working on this project.

Summary
- Primary language: Python (Streamlit app).
- How to run: create a virtualenv, install dependencies, then run Streamlit.

Quick start

1. Create and activate virtualenv

   On Windows:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   streamlit run app.py
   ```

2. Notes
- There are no automated tests or `tests/` directory in this repo.
- A `requirements.txt` file lists project dependencies.
- Ignore the Maven `mvn` tasks shown in the workspace — they appear unrelated to this Python project.

Key files
- [README.md](README.md) — user-facing project overview.
- [requirements.txt](requirements.txt) — Python dependencies.
- [app.py](app.py) — Streamlit application entry point.
- [spreadsheet.py](spreadsheet.py) — helper utilities.

Agent guidance (concise)
- Discover: read `README.md` and `requirements.txt` before making changes.
- Environment: use an isolated virtualenv; do not modify global Python packages.
- Run commands: prefer `streamlit run app.py` for manual checks.
- Tests: if adding tests, place them under `tests/` and document how to run them in `README.md`.
- Commits: keep changes focused and minimal; avoid large refactors without user consent.

Suggested follow-ups
- Create `.github/copilot-instructions.md` linking to this file for GitHub-hosted agent tooling.
- Add a small `tests/` scaffold and a `Makefile` or `task` for `pytest` to simplify automation.

If you want changes or more detail (example test scaffold, CI steps, or a dedicated skill), tell me which to create next.
