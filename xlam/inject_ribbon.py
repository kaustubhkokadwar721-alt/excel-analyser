"""inject_ribbon.py - inject the customUI14 ribbon part into a built .xlam.

VBComponents.Import cannot add a ribbon; the ribbon is an OOXML package part.
This rewrites the .xlam zip to add customUI/customUI14.xml, the package
relationship pointing to it, and (if missing) the xml content-type default.

Run:  python inject_ribbon.py [PerfDiag.xlam] [customUI14.xml]
"""

import os
import re
import sys
import shutil
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
XLAM = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "PerfDiag.xlam")
UIXML = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "customUI14.xml")

UI_PART = "customUI/customUI14.xml"
# Relationship type for the 2009/2010 custom UI part (customUI14.xml).
REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
REL_ID = "rIdPerfDiagUI"


def add_relationship(rels_xml: str) -> str:
    if UI_PART in rels_xml:
        return rels_xml  # already present
    rel = (f'<Relationship Id="{REL_ID}" Type="{REL_TYPE}" '
           f'Target="{UI_PART}"/>')
    return rels_xml.replace("</Relationships>", rel + "</Relationships>")


def ensure_xml_default(ct_xml: str) -> str:
    if re.search(r'<Default[^>]*Extension="xml"', ct_xml):
        return ct_xml
    default = '<Default Extension="xml" ContentType="application/xml"/>'
    # insert right after the opening <Types ...> tag
    return re.sub(r'(<Types[^>]*>)', r'\1' + default, ct_xml, count=1)


def main():
    if not os.path.exists(XLAM):
        print(f"ERROR: not found: {XLAM}")
        return 2
    with open(UIXML, "r", encoding="utf-8") as f:
        ui = f.read()

    tmp = XLAM + ".tmp"
    with zipfile.ZipFile(XLAM, "r") as zin:
        names = zin.namelist()
        rels = zin.read("_rels/.rels").decode("utf-8")
        cts = zin.read("[Content_Types].xml").decode("utf-8")
        new_rels = add_relationship(rels)
        new_cts = ensure_xml_default(cts)

        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in names:
                if n == "_rels/.rels":
                    zout.writestr(n, new_rels)
                elif n == "[Content_Types].xml":
                    zout.writestr(n, new_cts)
                elif n == UI_PART:
                    continue  # replaced below
                else:
                    zout.writestr(n, zin.read(n))
            zout.writestr(UI_PART, ui)

    shutil.move(tmp, XLAM)
    print(f"Injected ribbon into {os.path.basename(XLAM)}")
    print(f"  + {UI_PART}")
    print(f"  + relationship {REL_ID} -> {UI_PART}")
    print(f"  content-type xml default: {'ok' if 'Extension=\"xml\"' in new_cts else 'MISSING'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
