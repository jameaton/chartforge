"""Builds the ChartForge add-in from a file saved by PowerPoint.

PowerPoint has no UI for ribbon definitions, so this rewrites the package
with customUI/customUI.xml and the relationship that points to it. Given a
macro-enabled presentation (.pptm), it also switches the main part's content
type so the result is an add-in (.ppam), written next to the input.
Run: python3 tools/build_addin.py "ChartForge.pptm"
"""
import os
import re
import shutil
import sys
import tempfile
import zipfile

RIBBON = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "ribbon", "customUI.xml")
REL_TYPE = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
PPTM_MAIN = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
PPAM_MAIN = "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"
REL = f'<Relationship Id="afRibbon" Type="{REL_TYPE}" Target="customUI/customUI.xml"/>'


def main(path):
    if not os.path.isfile(path):
        sys.exit(f"Not found: {path}")
    with open(RIBBON, "rb") as f:
        ribbon = f.read()
    out = os.path.splitext(path)[0] + ".ppam"
    fd, tmp = tempfile.mkstemp(suffix=".ppam", dir=os.path.dirname(os.path.abspath(path)))
    os.close(fd)
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            if item.filename == "customUI/customUI.xml":
                continue
            data = src.read(item.filename)
            if item.filename == "_rels/.rels":
                rels = data.decode("utf-8")
                rels = re.sub(r'<Relationship [^>]*Type="' + re.escape(REL_TYPE) + r'"[^>]*/>', "", rels)
                rels = rels.replace("</Relationships>", REL + "</Relationships>")
                data = rels.encode("utf-8")
            if item.filename == "[Content_Types].xml":
                data = data.replace(PPTM_MAIN.encode(), PPAM_MAIN.encode())
            dst.writestr(item, data)
        dst.writestr("customUI/customUI.xml", ribbon)
    shutil.move(tmp, out)
    print(f"Add-in written to {out}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
