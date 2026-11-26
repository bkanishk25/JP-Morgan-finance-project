"""Recalculate every formula in an .xlsx workbook with LibreOffice (headless).

openpyxl writes formulas without cached values, so a freshly built workbook
reads as empty to anything that looks at values (pandas, openpyxl data_only).
This helper opens the file in LibreOffice, recalculates, and saves the values
back into a copy that sits next to the original.

Usage:  python src/recalc.py model/JPM_Financial_Model.xlsx
"""
import os
import shutil
import subprocess
import sys
import tempfile


def find_soffice():
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    mac = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if os.path.exists(mac):
        return mac
    win = r"C:\Program Files\LibreOffice\program\soffice.exe"
    if os.path.exists(win):
        return win
    sys.exit("LibreOffice not found. Install it from https://www.libreoffice.org/ and retry.")


def recalc(path, timeout=120):
    """Recalculate `path` in place. Returns the path."""
    path = os.path.abspath(path)
    soffice = find_soffice()
    with tempfile.TemporaryDirectory() as out_dir, tempfile.TemporaryDirectory() as profile:
        cmd = [
            soffice,
            f"-env:UserInstallation=file://{profile}",
            "--headless", "--norestore",
            "--convert-to", "xlsx:Calc MS Excel 2007 XML",
            "--outdir", out_dir, path,
        ]
        subprocess.run(cmd, check=True, capture_output=True, timeout=timeout)
        produced = os.path.join(out_dir, os.path.basename(path))
        if not os.path.exists(produced):
            raise RuntimeError("LibreOffice did not produce an output file")
        shutil.move(produced, path)
    return path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    print("recalculated:", recalc(sys.argv[1]))
