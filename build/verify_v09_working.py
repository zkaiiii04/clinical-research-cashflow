"""Check a built v0.9_working workbook against the v0.8 baseline.

Usage:
    python3 -I verify_v09_working.py <v0.8.xlsm> <v0.9_working.xlsm>

Checks:
1. The Power Query M code (customXml/item1.xml -> DataMashup -> Formulas/Section1.m)
   is identical, and every Power Query part (DataMashup, connections, query
   tables, query-table rels, PQ table definitions) is byte-identical. No
   exemptions.
2. Only the parts the builder is meant to touch differ; everything else is
   byte-identical.
3. Cross-part links (check_links), on both workbooks, because byte checks
   alone passed the broken first v0.9 (incident 2026-10-11):
   - every part parses, has a content type, and every relationship target
     exists;
   - every sheet named in a defined name exists, and each hidden
     ExternalData_n name sits on the sheet its formula points at;
   - every structured reference (Table[...]) names an existing table, and
     table names and ids are unique;
   - each query table's chain agrees: sheet -> table -> queryTable ->
     ExternalData_n name (same sheet, same range as the table) ->
     connection -> query, and the query's DataMashup FillTarget is that
     table's name. Every query that loads to a sheet has exactly one table.
4. The workbook opens with openpyxl (keep_vba=True); the new sheets, tables,
   names, KPI formulas and visibility are as the spec requires; the VBA
   project is untouched.
"""
import base64
import hashlib
import io
import posixpath
import re
import sys
import zipfile

import openpyxl
from lxml import etree

EXPECTED_CHANGED = {
    "[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels",
    "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet14.xml",
}
EXPECTED_REMOVED = {
    "xl/calcChain.xml", "xl/comments1.xml", "xl/drawings/vmlDrawing1.vml",
    "xl/worksheets/_rels/sheet1.xml.rels",
}
EXPECTED_ADDED = {
    "xl/worksheets/sheet18.xml", "xl/worksheets/sheet19.xml", "xl/worksheets/sheet20.xml",
    "xl/worksheets/_rels/sheet18.xml.rels", "xl/worksheets/_rels/sheet19.xml.rels",
    "xl/worksheets/_rels/sheet20.xml.rels", "xl/tables/table18.xml", "xl/tables/table19.xml",
}


def mashup(z):
    raw = z.read("customXml/item1.xml")
    txt = raw.decode("utf-16") if raw[:2] in (b"\xff\xfe", b"\xfe\xff") else raw.decode("utf-8")
    return base64.b64decode(re.search(r"<DataMashup[^>]*>(.*)</DataMashup>", txt, re.S).group(1))


def m_code(z):
    data = mashup(z)
    pkg_len = int.from_bytes(data[4:8], "little")
    inner = zipfile.ZipFile(io.BytesIO(data[8:8 + pkg_len]))
    return inner.read("Formulas/Section1.m").decode("utf-8-sig")


def fill_targets(z):
    """query name -> FillTarget table name, for queries that load to a sheet.

    DataMashup layout: version, then length-prefixed package, permissions and
    metadata blocks; the metadata block is version, length-prefixed XML, ...
    """
    data, pos = mashup(z), 4
    blocks = []
    for _ in range(3):
        n = int.from_bytes(data[pos:pos + 4], "little")
        blocks.append(data[pos + 4:pos + 4 + n])
        pos += 4 + n
    meta = blocks[2]
    xml = meta[8:8 + int.from_bytes(meta[4:8], "little")].decode("utf-8-sig")
    out = {}
    for item in re.findall(r"<Item>(.*?)</Item>", xml, re.S):
        path = re.search(r"<ItemPath>Section1/([^/<]+)</ItemPath>", item)
        if not path or 'Type="FillEnabled" Value="l1"' not in item:
            continue
        target = re.search(r'Type="FillTarget" Value="s([^"]*)"', item)
        out[path.group(1).replace("%20", " ")] = target.group(1) if target else None
    return out


NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "p": "http://schemas.openxmlformats.org/package/2006/relationships",
      "ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
SHEET_REF = re.compile(r"(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!")
STRUCT_REF = re.compile(r"([A-Za-z_\\][\w.]*)\[")


def rels(z, part):
    """Relationship Id -> (type suffix, target part) for a part."""
    folder, name = posixpath.split(part)
    rel_part = posixpath.join(folder, "_rels", name + ".rels")
    if rel_part not in z.namelist():
        return {}
    out = {}
    for r in etree.fromstring(z.read(rel_part)).findall("p:Relationship", NS):
        if r.get("TargetMode") == "External":
            continue
        target = r.get("Target")
        target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(
            posixpath.join(folder, target))
        out[r.get("Id")] = (r.get("Type").rsplit("/", 1)[-1], target)
    return out


def check_links(z):
    problems = []
    names = z.namelist()

    # Package integrity: parses, content types, relationship targets.
    ct = etree.fromstring(z.read("[Content_Types].xml"))
    defaults = {d.get("Extension").lower() for d in ct.findall("ct:Default", NS)}
    overrides = {o.get("PartName").lstrip("/") for o in ct.findall("ct:Override", NS)}
    for n in names:
        if n.endswith("/"):
            continue
        if n.endswith((".xml", ".rels", ".vml")):
            try:
                etree.fromstring(z.read(n))
            except etree.XMLSyntaxError as e:
                problems.append("%s does not parse: %s" % (n, e))
        if n not in overrides and n.rsplit(".", 1)[-1].lower() not in defaults:
            problems.append("No content type for " + n)
        if n.endswith(".rels"):
            src = n.replace("_rels/", "").rsplit(".rels", 1)[0]
            for _, target in rels(z, src).values():
                if target not in names:
                    problems.append("%s points at missing part %s" % (n, target))
    for o in overrides:
        if o not in names:
            problems.append("Content type for missing part " + o)

    wb = etree.fromstring(z.read("xl/workbook.xml"))
    wb_rels = rels(z, "xl/workbook.xml")
    sheets = [(s.get("name"), wb_rels[s.get("{%s}id" % NS["r"])][1])
              for s in wb.find("m:sheets", NS)]
    sheet_names = [s for s, _ in sheets]

    # Tables, by sheet.
    tables = {}       # table name (lower) -> (name, sheet, ref, part)
    table_ids = {}
    for sheet, sheet_part in sheets:
        for kind, part in rels(z, sheet_part).values():
            if kind != "table":
                continue
            t = etree.fromstring(z.read(part))
            name, ref = t.get("name"), t.get("ref")
            if t.get("displayName") != name:
                problems.append("Table %s has displayName %s" % (name, t.get("displayName")))
            if name.lower() in tables:
                problems.append("Duplicate table name " + name)
            if t.get("id") in table_ids:
                problems.append("Duplicate table id %s (%s, %s)" % (t.get("id"), name, table_ids[t.get("id")]))
            tables[name.lower()] = (name, sheet, ref, part)
            table_ids[t.get("id")] = name

    # Defined names: sheets exist; ExternalData_n on its own sheet.
    defined = {}
    dn = wb.find("m:definedNames", NS)
    formulas = []
    for d in (dn if dn is not None else []):
        text, name, local = d.text or "", d.get("name"), d.get("localSheetId")
        formulas.append(("defined name " + name, text))
        if local is not None and int(local) >= len(sheets):
            problems.append("Defined name %s has localSheetId %s (only %d sheets)" % (name, local, len(sheets)))
            continue
        refs = {(a.replace("''", "'") if a else b) for a, b in SHEET_REF.findall(text)}
        for sheet in refs:
            if sheet not in sheet_names:
                problems.append("Defined name %s refers to missing sheet %s: %s" % (name, sheet, text))
        if name.startswith("ExternalData_") and local is not None:
            owner = sheet_names[int(local)]
            if refs != {owner}:
                problems.append("%s (sheet %s) refers to %s" % (name, owner, text))
            defined[(owner, name)] = text.split("!", 1)[1].replace("$", "")

    # Structured references in defined names and cell formulas.
    for sheet, part in sheets:
        for f in etree.fromstring(z.read(part)).iter("{%s}f" % NS["m"]):
            if f.text:
                formulas.append((sheet, f.text))
    for where, text in formulas:
        for t in STRUCT_REF.findall(re.sub(r'"[^"]*"', '""', text)):
            if t.lower() not in tables and not t.startswith("_xl"):
                problems.append("%s refers to missing table %s" % (where, t))

    # Query tables: sheet -> table -> queryTable -> ExternalData_n -> connection -> FillTarget.
    conns = {}
    cx = etree.fromstring(z.read("xl/connections.xml"))
    for c in cx.findall("m:connection", NS):
        dbpr = c.find("m:dbPr", NS)
        cmd = dbpr.get("command", "") if dbpr is not None else ""
        q = re.match(r"SELECT \* FROM \[(.+)\]$", cmd)
        conns[c.get("id")] = q.group(1) if q else None
    targets = fill_targets(z)
    loaded = {}
    linked_qt = set()
    for name, sheet, ref, part in tables.values():
        for kind, qt_part in rels(z, part).values():
            if kind != "queryTable":
                continue
            linked_qt.add(qt_part)
            qt = etree.fromstring(z.read(qt_part))
            qt_name, conn = qt.get("name"), qt.get("connectionId")
            where = "%s (table %s on %s)" % (qt_part, name, sheet)
            if defined.get((sheet, qt_name)) != ref:
                problems.append("%s: hidden name %s is %r on that sheet, table range is %s"
                                % (where, qt_name, defined.get((sheet, qt_name)), ref))
            query = conns.get(conn)
            if query is None:
                problems.append("%s: connection %s missing or not a Power Query" % (where, conn))
                continue
            if targets.get(query, "") != name:
                problems.append("%s: query %s has FillTarget %r" % (where, query, targets.get(query)))
            loaded.setdefault(query, []).append(name)
    for qt_part in names:
        if qt_part.startswith("xl/queryTables/") and qt_part not in linked_qt:
            problems.append("Query table part not linked from any table: " + qt_part)
    for query, target in targets.items():
        if len(loaded.get(query, [])) != 1:
            problems.append("Query %s (FillTarget %s) loads to %s" % (query, target, loaded.get(query, [])))
    print("Links checked: %d sheets, %d tables, %d query tables, %d defined names"
          % (len(sheets), len(tables), len(linked_qt), 0 if dn is None else len(dn)))
    return problems


def is_pq_part(name, z):
    if name.startswith(("customXml/", "xl/queryTables/")) or name == "xl/connections.xml":
        return True
    if re.match(r"xl/tables/(_rels/)?table\d+\.xml(\.rels)?$", name):
        return b'tableType="queryTable"' in z.read(name) or name.endswith(".rels")
    return False


def main(old_path, new_path):
    problems = []
    zo, zn = zipfile.ZipFile(old_path), zipfile.ZipFile(new_path)
    old, new = set(zo.namelist()), set(zn.namelist())

    mo, mn = m_code(zo), m_code(zn)
    print("M code sha256 v0.8 :", hashlib.sha256(mo.encode()).hexdigest())
    print("M code sha256 v0.9 :", hashlib.sha256(mn.encode()).hexdigest())
    if mo != mn:
        problems.append("Power Query M code differs")
    queries = re.findall(r'shared\s+(#?"[^"]+"|\w+)\s*=', mn)
    print("Queries (%d): %s" % (len(queries), ", ".join(queries)))

    pq_parts = sorted(n for n in old if is_pq_part(n, zo))
    for n in pq_parts:
        if n not in new:
            problems.append("Power Query part missing: " + n)
        elif zo.read(n) != zn.read(n):
            problems.append("Power Query part changed: " + n)
    print("Power Query parts checked: %d (all must be byte-identical)" % len(pq_parts))

    for label, z in (("v0.8", zo), ("v0.9", zn)):
        app = re.search(rb"<Application>([^<]*)</Application>", z.read("docProps/app.xml"))
        print("%s last saved by: %s" % (label, app.group(1).decode() if app else "?"))
        found = check_links(z)
        problems += ["%s: %s" % (label, p) for p in found]

    if old - new != EXPECTED_REMOVED:
        problems.append("Unexpected removed parts: %s" % sorted((old - new) ^ EXPECTED_REMOVED))
    if new - old != EXPECTED_ADDED:
        problems.append("Unexpected added parts: %s" % sorted((new - old) ^ EXPECTED_ADDED))
    changed = {n for n in old & new if zo.read(n) != zn.read(n)}
    if changed - EXPECTED_CHANGED:
        problems.append("Unexpected changed parts: %s" % sorted(changed - EXPECTED_CHANGED))
    print("Changed parts: %s" % sorted(changed))
    if zo.read("xl/vbaProject.bin") != zn.read("xl/vbaProject.bin"):
        problems.append("vbaProject.bin changed")

    wb = openpyxl.load_workbook(new_path, keep_vba=True)
    states = {ws.title: ws.sheet_state for ws in wb.worksheets}
    for name, state in [("Main_Dashboard", "visible"), ("Legacy_Timeline", "hidden"),
                        ("Dashboard_Settings", "visible"), ("ProjectionStore", "veryHidden")]:
        if states.get(name) != state:
            problems.append("%s state is %s, expected %s" % (name, states.get(name), state))

    tables = {t: (ws.title, ws.tables[t].ref) for ws in wb.worksheets for t in ws.tables}
    expect_tables = {
        "tblManualProjection": ("Legacy_Timeline", "B59:G60"),
        "tblPatientReimbursement": ("Legacy_Timeline", "B66:H67"),
        "tblManpowerChargeTo": ("Dashboard_Settings", "A4:E5"),
        "tblProjectionStore": ("ProjectionStore", "A1:K2"),
        "qManpower_SelectedTest": ("PQManpower_SelectedTest", "A1:R13"),
    }
    for t, where in expect_tables.items():
        if tables.get(t) != where:
            problems.append("Table %s at %s, expected %s" % (t, tables.get(t), where))

    names = {n: wb.defined_names[n].attr_text for n in wb.defined_names}
    if names.get("TL_Anchor") != "Main_Dashboard!$B$21":
        problems.append("TL_Anchor is %r" % names.get("TL_Anchor"))
    for n in ("TL_TotalBudget", "TL_TotalActual", "TL_TotalProjected"):
        if names.get(n) != '""':
            problems.append("%s is %r" % (n, names.get(n)))

    d = wb["Main_Dashboard"]
    expect = {"C14": "=TL_TotalBudget", "F14": "=TL_TotalActual", "F17": "=TL_TotalProjected",
              "C17": '=IFERROR(F14/C14,"")', "C19": '=IFERROR(F17/C14,"")'}
    for ref, f in expect.items():
        if d[ref].value != f:
            problems.append("Main_Dashboard!%s = %r" % (ref, d[ref].value))
    for ref in ("C7", "C8"):
        if "tblSelectedProjectMapping" not in str(d[ref].value):
            problems.append("Main_Dashboard!%s not repointed" % ref)
    if "tblActualsSelectedTest[Posting Month]" not in str(d["C9"].value):
        problems.append("Main_Dashboard!C9 not repointed")
    for ref in ("B5", "C5", "L14"):
        if d[ref].value is not None:
            problems.append("Main_Dashboard!%s should be empty" % ref)
    if d.max_row > 20:
        problems.append("Main_Dashboard has content below row 20 (max_row %d)" % d.max_row)
    dvs = {str(v.sqref): v.formula1 for v in d.data_validations.dataValidation}
    if "C5" in dvs or dvs.get("C6") != '"Financial Year,Calendar Year"':
        problems.append("Main_Dashboard validations: %s" % dvs)
    for ref in ("E6", "F6", "G6"):
        if d[ref].value is None:
            problems.append("Selector %s missing" % ref)

    leg = wb["Legacy_Timeline"]
    if not str(leg["H25"].value).startswith("=IF(INDEX($21:$21"):
        problems.append("Legacy_Timeline timeline formulas missing")
    if leg["C5"].value is None:
        problems.append("Legacy_Timeline lost its C5 driver")
    if wb["PQManpower_SelectedTest"]["P15"].value is not None:
        problems.append("Stray SUM still in PQManpower_SelectedTest!P15")

    for p in problems:
        print("PROBLEM:", p)
    print("RESULT:", "PASS" if not problems else "FAIL (%d)" % len(problems))
    return 0 if not problems else 1


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    sys.exit(main(sys.argv[1], sys.argv[2]))
