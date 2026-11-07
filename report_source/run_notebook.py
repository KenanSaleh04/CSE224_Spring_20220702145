"""Execute the submission notebook top to bottom in a fresh kernel and save its outputs.

Usage: python report_source/run_notebook.py
The kernel's working directory is report_source/, so the relative outputs/ folder
(figures + results.json used by build_report.py) is created there.
"""
from pathlib import Path
import nbformat
from nbconvert.preprocessors import ExecutePreprocessor

ROOT = Path(__file__).resolve().parents[1]
NB = ROOT / "submission" / "CSE464_Case_Study_1.ipynb"
nb = nbformat.read(NB, as_version=4)
ExecutePreprocessor(timeout=1800, kernel_name="python3").preprocess(nb, {"metadata": {"path": str(ROOT / "report_source")}})
nbformat.write(nb, NB)
print("Executed and saved", NB)
