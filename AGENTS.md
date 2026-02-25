# AGENTS.md

## Cursor Cloud specific instructions

### Project overview

This is the **Excel Data Merger/Aggregator Tool** (Excel数据汇总工具) — a single-file Python/Flask web application that lets users upload multiple Excel files, configure column mappings, merge them, and download the result. The frontend is a single Jinja2-rendered `templates/index.html` with no build step.

### Branch layout

The application code lives on the `excelmergertool` branch. The `main` branch contains only `README.md`.

### Running the dev server

```bash
source /workspace/venv/bin/activate
python app.py
```

The Flask server starts on `http://127.0.0.1:5000` with `debug=True` in development mode. It will attempt to auto-open a browser (harmless dbus errors appear in headless environments — ignore them).

### Caveats

- **No automated tests exist** in this repository. There is no test suite, linter config, or CI test step. Validation is manual via the API or UI.
- **No linter config** is present (no `flake8`, `pylint`, `mypy`, or similar). You can install and run `flake8` ad-hoc if needed.
- `pyinstaller` is listed in `requirements.txt` for building a Windows `.exe` — it is not needed for development.
- Uploaded files are stored in `uploads/` and merged output in `output/` within the project root. These directories are created automatically and are gitignored.
- The app binds to `127.0.0.1:5000` only (not `0.0.0.0`).
