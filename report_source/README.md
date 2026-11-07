# CSE464 Case Study 1 – report source

Submission files are in `../submission/`:

- `CSE464_Case_Study_1.pdf` – the report (built from `CSE464_Case_Study_1.docx`)
- `CSE464_Case_Study_1.ipynb` – the executed analysis notebook

## Files here

| File | Purpose |
|---|---|
| `CSE464_Case_Study_1.docx` | Editable report; open it in Word to change wording |
| `build_notebook.py` | Writes the notebook (all cells) to `../submission/` |
| `run_notebook.py` | Executes the notebook top to bottom and saves its outputs; creates `outputs/` |
| `build_report.py` | Builds the DOCX and PDF; every number comes from `outputs/results.json` |
| `assets/yeditepe_header.png` | University header taken from the course template |
| `outputs/` | Figures, `results.json`, monthly and country CSVs written by the notebook |

## Rebuild everything (local)

```bash
pip install pandas numpy matplotlib ucimlrepo openpyxl nbformat nbconvert ipykernel python-docx
python report_source/build_notebook.py   # overwrites the notebook without outputs
python report_source/run_notebook.py     # runs it and saves the outputs
python report_source/build_report.py     # needs LibreOffice (soffice) for the PDF
```

If you edit the DOCX by hand, export it to PDF from Word yourself. Running `build_report.py` again overwrites your manual edits.

## Run in Google Colab

1. Go to colab.research.google.com, choose **File ▸ Upload notebook**, and select `submission/CSE464_Case_Study_1.ipynb`.
2. Choose **Runtime ▸ Run all**. The first cell installs `ucimlrepo`, and Step 1 downloads the data from UCI. No Drive mount is needed.
3. Check that every cell finished without an error. Then choose **File ▸ Save**, followed by **File ▸ Download ▸ Download .ipynb**.
