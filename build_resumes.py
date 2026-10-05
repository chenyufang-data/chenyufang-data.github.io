"""Make the web copies of the three resumes for cyfang.org: the same layout, with the phone number removed.

Run after changing a resume:   python build_resumes.py

The originals in JobNavigator/chenyufang are only read. The number is removed from a copy's XML (the package is
rewritten part by part, in order), and Word only opens the copy and exports it to PDF. Word must not edit the
document itself or write into this folder directly: in a hidden instance either can hang it.
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

SOURCE = Path.home() / "Documents" / "JobNavigator" / "chenyufang"
# any US phone number in the header line, with the separator after it (no number is stored here)
PHONE = re.compile(rb"\(\d{3}\) \d{3}-\d{4} \| ")
NAMES = {"MLE": "ML_Engineer", "AIE": "AI_Engineer", "ASci": "Applied_Scientist"}
HERE = Path(__file__).resolve().parent
OUT = HERE / "resume"
WORK = HERE / "_build"

EXPORT = r"""
$ErrorActionPreference = 'Stop'
$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
try {
  foreach ($pair in $env:CYFANG_JOBS.Split(';')) {
    $src, $pdf = $pair.Split('|')
    $d = $w.Documents.Open($src, $false, $true, $false)
    $d.ExportAsFixedFormat($pdf, 17)
    $d.Close(0)
    "exported $pdf"
  }
} finally { $w.Quit() }
"""


def strip_phone(src: Path, dst: Path) -> None:
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/document.xml":
                data, n = PHONE.subn(b"", data, count=1)
                if n != 1:
                    raise SystemExit(f"phone number not found in {src.name}")
            zout.writestr(info, data)


def kill_hidden_word() -> None:
    subprocess.run(["powershell", "-NoProfile", "-Command",
                    "Get-CimInstance Win32_Process -Filter \"Name='WINWORD.EXE'\" | "
                    "Where-Object { $_.CommandLine -like '*Automation*' } | "
                    "ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"], check=False)


def main() -> int:
    OUT.mkdir(exist_ok=True)
    shutil.rmtree(WORK, ignore_errors=True)
    WORK.mkdir()
    tmp = Path(tempfile.gettempdir())
    jobs = []
    for key, name in NAMES.items():
        copy = WORK / f"web_{key}.docx"
        strip_phone(SOURCE / f"Chenyu_Fang_resume_{key}.docx", copy)
        jobs.append((copy, tmp / f"cyfang_{name}.pdf", OUT / f"Chenyu_Fang_Resume_{name}.pdf"))
    env = {**os.environ, "CYFANG_JOBS": ";".join(f"{c}|{t}" for c, t, _ in jobs)}
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-Command", EXPORT], env=env, capture_output=True,
                           text=True, timeout=180)
    except subprocess.TimeoutExpired:
        kill_hidden_word()
        print("Word did not finish in 3 minutes; stopped it. Nothing was changed.", file=sys.stderr)
        return 1
    if r.returncode != 0:
        print(r.stdout, r.stderr, file=sys.stderr)
        return 1
    for _, t, final in jobs:
        shutil.move(str(t), final)
        print(final.relative_to(HERE))
    shutil.rmtree(WORK, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
