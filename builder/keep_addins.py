# -*- coding: utf-8 -*-
"""
Put Office add-in bindings back into a workbook after openpyxl has saved it.

openpyxl silently drops the xl/webextensions parts on save. Donabay.xlsx
carries two of them, one being the Claude for Excel add-in (property
claude.fileId); without it the add-in loses its link to the workbook.

    python3 builder/keep_addins.py ORIGINAL.xlsx EDITED.xlsx

EDITED.xlsx is rewritten in place with ORIGINAL's add-in parts, their package
relationship and their content types. Running it twice does nothing extra.
"""
import os, re, sys, zipfile

TASKPANES_REL = "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"


def keep_addins(original, edited):
    with zipfile.ZipFile(original) as src:
        parts = {n: src.read(n) for n in src.namelist() if n.startswith("xl/webextensions/")}
        overrides = [m for m in re.findall(r"<Override[^>]*/>", src.read("[Content_Types].xml").decode())
                     if "/xl/webextensions/" in m]
    if not parts:
        return 0
    with zipfile.ZipFile(edited) as dst:
        if any(n.startswith("xl/webextensions/") for n in dst.namelist()):
            return 0
        items = [(info, dst.read(info.filename)) for info in dst.infolist()]

    tmp = edited + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for info, data in items:
            if info.filename == "[Content_Types].xml":
                data = data.decode().replace("</Types>", "".join(overrides) + "</Types>").encode()
            elif info.filename == "_rels/.rels":
                s = data.decode()
                next_id = max(int(i) for i in re.findall(r'Id="rId(\d+)"', s)) + 1
                s = s.replace("</Relationships>",
                              f'<Relationship Id="rId{next_id}" Type="{TASKPANES_REL}" '
                              f'Target="xl/webextensions/taskpanes.xml"/></Relationships>')
                data = s.encode()
            out.writestr(info, data)
        for name, data in parts.items():
            out.writestr(name, data)
    os.replace(tmp, edited)
    return len(parts)


if __name__ == "__main__":
    n = keep_addins(sys.argv[1], sys.argv[2])
    print(f"restored {n} add-in part(s)" if n else "nothing to restore")
