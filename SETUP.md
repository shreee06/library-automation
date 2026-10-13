# Running this in VS Code

## 1. Open the folder
`File > Open Folder...` and pick this project folder.

## 2. Create a virtual environment (recommended)
Open a terminal in VS Code (`` Ctrl+` ``):
```bash
python -m venv venv
# Windows:
venv\Scripts\activate
# Mac/Linux:
source venv/bin/activate
```
VS Code may prompt "Select this environment for the workspace?" — click Yes.
(If using the Python extension, you can also pick the interpreter via
`Ctrl+Shift+P` -> "Python: Select Interpreter" -> the `venv` one.)

## 3. Install dependencies
```bash
pip install -r requirements.txt
```

## 4. Add your Pangram API key
```bash
cp .env.example .env
```
Open `.env` and paste your real key in place of `your_pangram_api_key_here`.
(`.env` should NOT be committed to git if you set this up as a repo later —
add it to `.gitignore`.)

## 5. Add a test PDF
Create an `input/` folder (if it doesn't exist) and drop in a PDF named like:
```
21CS045_John_Doe_CSE.pdf
```
`<RegNo>_<Student Name>_<Department>.pdf` — this stands in for the email
metadata until you have library inbox access.

## 6. Run it
```bash
python main.py
```
Check `output/master_log.xlsx` for the logged result. Processed PDFs move to
`input/processed/`; anything that fails parsing or the Pangram check moves
to `input/needs_review/` and is still logged with an error status, so
nothing silently disappears.

## 7. Debugging in VS Code
Each file (`filename_parser.py`, `pdf_extractor.py`, `pangram_check.py`,
`excel_logger.py`) can also be run on its own for quick testing — each has
a `if __name__ == "__main__":` block at the bottom. Right-click any of them
in the editor and choose "Run Python File in Terminal", or set breakpoints
and use the Run/Debug panel (`F5`) to step through `main.py`.

## First thing to check once you have a real API key
`pangram_check.py`'s field-name guesses (`ai_likelihood`, `tier`, etc.) may
not exactly match what Pangram actually returns — the public docs we had
were incomplete. Run:
```bash
python pangram_check.py "This is a short test sentence to check."
```
and look at the printed "Raw predict() response" — adjust the key names in
`_safe_get(...)` calls in `pangram_check.py` to match. That's the only part
likely to need a tweak once you're against the live API.
