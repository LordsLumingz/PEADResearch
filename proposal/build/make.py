"""Two-pass build: render, read page numbers from the PDF, rebuild with a filled TOC.

Usage: python make.py OUTDIR
Writes OUTDIR/Proposal_PEAD_LQ45_Azim.docx and .pdf.
"""
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUTDIR = Path(sys.argv[1]).resolve()
NAME = "Proposal_PEAD_LQ45_Azim"
FIG = HERE / "fig"

HEADINGS = [
    "1.1 Research Background", "1.2 Problem Definition",
    "1.3 Research Questions and Research Objectives",
    "1.4 Scope and Significance of the Study", "1.5 Research Novelty",
    "2.1 Literature Synthesis", "2.2 Theoretical Framework and Conceptual Model",
    "2.3 Hypotheses Development", "3.1 Research Design",
    "3.2 Operational Definition and Measurement", "3.3 Population and Sampling Design",
    "3.4 Data Collection Methods and Bias Reduction", "3.5 Data Analysis Plan",
    "3.6 Research Ethics and Thesis Timeline",
]
CHAPTERS = {"BAB I PENDAHULUAN": "BAB I",
            "BAB II LITERATURE REVIEW AND THEORETICAL FRAMEWORK": "BAB II",
            "BAB III RESEARCH METHODOLOGY": "BAB III"}
TABLES = ["1.1", "1.2", "1.3", "1.4", "2.1", "2.2", "2.3", "2.4", "3.1", "3.2", "3.3", "3.4", "3.5",
          "3.6", "3.7", "3.8", "3.9"]
FIGS = ["2.1", "3.1", "3.2"]


def roman(n):
    vals = [(10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out


def norm(s):
    return re.sub(r"\s+", " ", s).strip()


def build(docx, pages_json=None):
    cmd = [sys.executable, str(HERE / "build.py"), str(docx), str(FIG)]
    if pages_json:
        cmd.append(str(pages_json))
    subprocess.run(cmd, check=True, cwd=HERE)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir",
                    str(docx.parent), str(docx)], check=True, capture_output=True)
    return docx.with_suffix(".pdf")


def page_texts(pdf):
    n = int(re.search(r"Pages:\s+(\d+)",
                      subprocess.run(["pdfinfo", str(pdf)], capture_output=True,
                                     text=True).stdout).group(1))
    texts = []
    for i in range(1, n + 1):
        t = subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), "-layout", str(pdf), "-"],
                           capture_output=True, text=True).stdout
        texts.append(t)
    return texts


def locate(pdf):
    texts = page_texts(pdf)
    lines = [[norm(l) for l in t.splitlines()] for t in texts]
    # body starts at the page whose lines contain exactly "BAB I"
    body = next(i for i, ls in enumerate(lines) if "BAB I" in ls)
    front = 1  # page index 1 (second page) is roman i
    pages = {}

    def body_no(i):
        return str(i - body + 1)

    for key, label in CHAPTERS.items():
        i = next(i for i in range(body, len(lines)) if label in lines[i])
        pages[key] = body_no(i)
    for h in HEADINGS:
        i = next(i for i in range(body, len(lines)) if any(l.startswith(h) for l in lines[i]))
        pages[h] = body_no(i)
    for t in TABLES:
        i = next(i for i in range(body, len(lines))
                 if any(l.startswith(f"Table {t}. ") for l in lines[i]))
        pages[f"Table {t}"] = body_no(i)
    for f in FIGS:
        i = next(i for i in range(body, len(lines))
                 if any(l.startswith(f"Figure {f}. ") for l in lines[i]))
        pages[f"Figure {f}"] = body_no(i)
    i = next(i for i in range(body, len(lines)) if "DAFTAR PUSTAKA" in lines[i])
    pages["DAFTAR PUSTAKA"] = body_no(i)
    for key in ["DAFTAR TABEL", "DAFTAR GAMBAR", "DAFTAR ISTILAH"]:
        i = next(i for i in range(2, body) if key in lines[i])
        pages[key] = roman(i - front + 1)
    return pages


with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    pdf = build(tmp / "pass.docx")
    pages = locate(pdf)
    for _ in range(3):
        pj = tmp / "pages.json"
        pj.write_text(json.dumps(pages))
        OUTDIR.mkdir(parents=True, exist_ok=True)
        final_pdf = build(OUTDIR / f"{NAME}.docx", pj)
        new = locate(final_pdf)
        if new == pages:
            break
        pages = new
print(json.dumps(pages, indent=1))
print("done:", OUTDIR / f"{NAME}.docx")
