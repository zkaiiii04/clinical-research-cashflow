"""Build Clinical_Research_Cash_Flow_SharePoint_PQ_Ready_v0.9_working.xlsm from v0.8.

Why this edits the package XML instead of using openpyxl
---------------------------------------------------------
An openpyxl load/save of v0.8 drops customXml/item1.xml (the Power Query
DataMashup), xl/connections.xml, every xl/queryTables part and
xl/metadata.xml (dynamic arrays). That would destroy the workbook's queries.
So this script copies every part of the v0.8 package unchanged and edits only
the parts listed below. All Power Query parts stay byte-identical, and
verify_v09_working.py checks that.

What it changes (spec §4, §7, §8, §12; HANDOVER §4 steps 1-6)
-------------------------------------------------------------
* Main_Dashboard (sheet1.xml):
  - removes the old timeline and input tables (rows 21-71); the macro builds
    the new grid from B21 (TL_Anchor)
  - retires the C5 Project Display Name selector (B5/C5 and its validation)
  - removes the H3/J3 links to the retired input tables and the hidden L14
    helper
  - repoints C7, C8, C9, C14, F14, F17, C17 and C19 (spec §12)
  - adds a list validation to C6 (Financial Year / Calendar Year)
  - writes the "not loaded yet" notice in E7, below the selector
  - drops the old column outline and the two obsolete notes (G3, C5)
* Legacy_Timeline (new, hidden): an exact copy of the v0.8 Main_Dashboard,
  still driven by its own C5, carrying tblManualProjection and
  tblPatientReimbursement. It is for UAT comparison and is deleted after
  sign-off.
* Dashboard_Settings (new): tblManpowerChargeTo, with a list validation on
  Charge To.
* ProjectionStore (new, very hidden): tblProjectionStore, plus the grid
  layout area the macros write on each build.
* PQManpower_SelectedTest: removes the stray =SUM(...) in row 15, and (at the
  user's request, 11 Oct 2026) renames the sheet PQ_Manpower_SelectedTest and
  its table tblManpowerSelectedTest, following PQ_Budget_SelectedTest /
  tblBudgetSelectedTest. The Power Query query keeps its name
  (qManpower_SelectedTest), and no query reads the table, so only the table's
  name/displayName and the sheet name change.
* workbook.xml: new sheets and the TL_* names; full recalculation on load.
* calcChain.xml is removed; Excel rebuilds it on the next save.

Usage:
    python3 -I build_v09_working.py <v0.8.xlsm> <v0.9_working.xlsm>
"""
import re
import sys
import uuid
import zipfile
from copy import deepcopy

from lxml import etree

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_WS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
REL_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
CT_WS = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
CT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
M = "{%s}" % NS_MAIN
R = "{%s}" % NS_R

DASH = "xl/worksheets/sheet1.xml"
MANPOWER_SHEET = "xl/worksheets/sheet14.xml"
LEGACY = "xl/worksheets/sheet18.xml"
SETTINGS = "xl/worksheets/sheet19.xml"
STORE = "xl/worksheets/sheet20.xml"
T_CHARGE = "xl/tables/table18.xml"
T_STORE = "xl/tables/table19.xml"
T_MANPOWER = "xl/tables/table15.xml"
MANPOWER_RENAME = {"table": ("qManpower_SelectedTest", "tblManpowerSelectedTest"),
                   "sheet": ("PQManpower_SelectedTest", "PQ_Manpower_SelectedTest")}

STORE_COLUMNS = [
    "Source Project", "Main IO", "IO Code", "Expense Category", "Budget Code",
    "Line Key", "Budget Description", "Month", "Amount", "Saved By", "Saved At",
]
CHARGE_COLUMNS = ["Source Project", "Main IO", "Rate Category", "IO Code", "Charge To"]

# Spec §12. Names are defined as "" so the KPI cards are blank until the
# first Load Project redefines them to point at the grid.
KPI_FORMULAS = {
    "C7": 'IFERROR(INDEX(tblSelectedProjectMapping[Project Start],'
          'MATCH("Yes",tblSelectedProjectMapping[Main IO],0)),"")',
    "C8": 'IFERROR(INDEX(tblSelectedProjectMapping[Project End],'
          'MATCH("Yes",tblSelectedProjectMapping[Main IO],0)),"")',
    "C9": 'IFERROR(IF(COUNT(tblActualsSelectedTest[Posting Month])=0,"",'
          'MAX(tblActualsSelectedTest[Posting Month])),"")',
    "C14": "TL_TotalBudget",
    "F14": "TL_TotalActual",
    "F17": "TL_TotalProjected",
    "C17": 'IFERROR(F14/C14,"")',
    "C19": 'IFERROR(F17/C14,"")',
}
NEW_NAMES = [
    ("TL_Anchor", "Main_Dashboard!$B$21"),
    ("TL_TotalActual", '""'),
    ("TL_TotalBudget", '""'),
    ("TL_TotalProjected", '""'),
]
NOT_LOADED = "Dashboard not loaded yet. Choose E6/F6, then click Load Project."
SUBTITLE = ("v0.9 working | Project chosen by Source Project (E6) + Main IO (F6). "
            "Click Load Project to build the timeline.")
LEGACY_BANNER = ("LEGACY v0.8 dashboard - hidden copy for UAT comparison only "
                 "(driven by its own C5). Delete after UAT sign-off.")


def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    return m.group(1), int(m.group(2))


def num_to_col(n):
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def parse(data):
    return etree.fromstring(data, etree.XMLParser(remove_blank_text=False))


def dump(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def new_uid():
    return "{%s}" % str(uuid.uuid4()).upper()


class SharedStrings:
    """Appends new strings to xl/sharedStrings.xml and returns their index."""

    def __init__(self, data):
        self.root = parse(data)
        self.items = self.root.findall(M + "si")
        self.index = {}
        for i, si in enumerate(self.items):
            t = si.find(M + "t")
            if t is not None and len(si) == 1:
                self.index.setdefault(t.text or "", i)

    def add(self, text):
        if text in self.index:
            return self.index[text]
        si = etree.SubElement(self.root, M + "si")
        t = etree.SubElement(si, M + "t")
        t.text = text
        if text != text.strip():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        i = len(self.items)
        self.items.append(si)
        self.index[text] = i
        return i

    def bytes(self):
        self.root.set("uniqueCount", str(len(self.items)))
        count = int(self.root.get("count", "0"))
        self.root.set("count", str(max(count, len(self.items))))
        return dump(self.root)


def find_cell(sheet_data, ref):
    _, row = split_ref(ref)
    for r in sheet_data.findall(M + "row"):
        if int(r.get("r")) == row:
            for c in r.findall(M + "c"):
                if c.get("r") == ref:
                    return r, c
            return r, None
    return None, None


def ensure_row(sheet_data, row):
    rows = sheet_data.findall(M + "row")
    for r in rows:
        if int(r.get("r")) == row:
            return r
    new = etree.Element(M + "row", r=str(row))
    for i, r in enumerate(rows):
        if int(r.get("r")) > row:
            r.addprevious(new)
            return new
    sheet_data.append(new)
    return new


def put_cell(sheet_data, ref, style=None):
    """Return the <c> for ref, creating it in column order if needed, emptied."""
    col, row = split_ref(ref)
    r = ensure_row(sheet_data, row)
    r.attrib.pop("spans", None)
    for c in r.findall(M + "c"):
        if c.get("r") == ref:
            for child in list(c):
                c.remove(child)
            c.attrib.pop("t", None)
            c.attrib.pop("cm", None)
            if style is not None:
                c.set("s", str(style))
            return c
    new = etree.Element(M + "c", r=ref)
    if style is not None:
        new.set("s", str(style))
    for c in r.findall(M + "c"):
        if col_to_num(split_ref(c.get("r"))[0]) > col_to_num(col):
            c.addprevious(new)
            return new
    r.append(new)
    return new


def set_string(sheet_data, sst, ref, text, style=None):
    c = put_cell(sheet_data, ref, style)
    c.set("t", "s")
    etree.SubElement(c, M + "v").text = str(sst.add(text))


def set_formula(sheet_data, ref, formula):
    _, c = find_cell(sheet_data, ref)
    style = c.get("s") if c is not None else None
    c = put_cell(sheet_data, ref, style)
    etree.SubElement(c, M + "f").text = formula


def remove_cell(sheet_data, ref):
    r, c = find_cell(sheet_data, ref)
    if c is not None:
        r.remove(c)
        r.attrib.pop("spans", None)


def check_no_shared_formula_from_removed_rows(sheet_data, last_kept_row):
    """A kept cell must not depend on a shared-formula master in a removed row."""
    masters = {}
    for r in sheet_data.findall(M + "row"):
        for c in r.findall(M + "c"):
            f = c.find(M + "f")
            if f is not None and f.get("t") == "shared" and f.text:
                masters[f.get("si")] = int(r.get("r"))
    for r in sheet_data.findall(M + "row"):
        if int(r.get("r")) > last_kept_row:
            continue
        for c in r.findall(M + "c"):
            f = c.find(M + "f")
            if f is not None and f.get("t") == "shared":
                if masters.get(f.get("si"), 0) > last_kept_row:
                    raise SystemExit("Shared formula in %s depends on a removed row" % c.get("r"))


def edit_dashboard(data, sst):
    root = parse(data)
    sd = root.find(M + "sheetData")
    check_no_shared_formula_from_removed_rows(sd, 20)

    for r in sd.findall(M + "row"):
        if int(r.get("r")) >= 21:
            sd.remove(r)

    for ref in ("B5", "C5", "H3", "J3", "L14"):
        remove_cell(sd, ref)
    for ref, formula in KPI_FORMULAS.items():
        set_formula(sd, ref, formula)
    _, b2 = find_cell(sd, "B2")
    set_string(sd, sst, "B2", SUBTITLE, b2.get("s"))
    set_string(sd, sst, "E7", NOT_LOADED)

    root.find(M + "dimension").set("ref", "B1:BN20")

    view = root.find(M + "sheetViews").find(M + "sheetView")
    view.set("topLeftCell", "A1")
    view.set("tabSelected", "1")
    sel = view.find(M + "selection")
    if sel is not None:
        sel.set("activeCell", "E6")
        sel.set("sqref", "E6")

    fmt = root.find(M + "sheetFormatPr")
    fmt.attrib.pop("outlineLevelRow", None)
    fmt.attrib.pop("outlineLevelCol", None)

    cols = root.find(M + "cols")
    for col in cols.findall(M + "col"):
        if int(col.get("min")) >= 8:
            cols.remove(col)

    merges = root.find(M + "mergeCells")
    for mc in merges.findall(M + "mergeCell"):
        if int(re.findall(r"\d+", mc.get("ref"))[0]) >= 21:
            merges.remove(mc)
    merges.set("count", str(len(merges)))

    dvs = root.find(M + "dataValidations")
    for dv in dvs.findall(M + "dataValidation"):
        if dv.get("sqref") == "C5":
            dvs.remove(dv)
    basis = etree.SubElement(dvs, M + "dataValidation", type="list", allowBlank="1",
                             showInputMessage="1", showErrorMessage="1", sqref="C6")
    etree.SubElement(basis, M + "formula1").text = '"Financial Year,Calendar Year"'
    dvs.set("count", str(len(dvs)))

    links = root.find(M + "hyperlinks")
    for h in links.findall(M + "hyperlink"):
        if int(re.findall(r"\d+", h.get("ref"))[0]) >= 21:
            links.remove(h)

    for tag in ("legacyDrawing", "tableParts"):
        el = root.find(M + tag)
        if el is not None:
            root.remove(el)
    return dump(root)


def make_legacy(original_dashboard, sst):
    root = parse(original_dashboard)
    root.set("{http://schemas.microsoft.com/office/spreadsheetml/2014/revision}uid", new_uid())
    pr = root.find(M + "sheetPr")
    if pr is not None:
        pr.attrib.pop("codeName", None)
    view = root.find(M + "sheetViews").find(M + "sheetView")
    view.attrib.pop("tabSelected", None)
    view.set("topLeftCell", "A1")
    sd = root.find(M + "sheetData")
    _, b1 = find_cell(sd, "B1")
    set_string(sd, sst, "B1", LEGACY_BANNER, b1.get("s"))
    for h in root.find(M + "hyperlinks").findall(M + "hyperlink"):
        if h.get("location") == "'Main_Dashboard'!A1":
            h.set("location", "'Legacy_Timeline'!A1")
    # The two notes (G3 placeholder, C5 help) are obsolete; drop the drawing.
    ld = root.find(M + "legacyDrawing")
    root.remove(ld)
    parts = root.find(M + "tableParts")
    for i, tp in enumerate(parts.findall(M + "tablePart"), start=1):
        tp.set(R + "id", "rId%d" % i)
    return dump(root)


def sheet_xml(cells, cols, table_rid, dv=None, tab_color=None):
    """Minimal worksheet. cells: {(row, col): ("s", sst_index) | ("n", number)}."""
    ws = etree.Element(M + "worksheet", nsmap={None: NS_MAIN, "r": NS_R})
    pr = etree.SubElement(ws, M + "sheetPr")
    if tab_color:
        etree.SubElement(pr, M + "tabColor", rgb=tab_color)
    last_row = max(r for r, _ in cells)
    last_col = max(c for _, c in cells)
    etree.SubElement(ws, M + "dimension", ref="A1:%s%d" % (num_to_col(last_col), last_row))
    views = etree.SubElement(ws, M + "sheetViews")
    etree.SubElement(views, M + "sheetView", workbookViewId="0")
    etree.SubElement(ws, M + "sheetFormatPr", defaultRowHeight="15")
    cols_el = etree.SubElement(ws, M + "cols")
    for i, width in enumerate(cols, start=1):
        etree.SubElement(cols_el, M + "col", min=str(i), max=str(i), width=str(width), customWidth="1")
    sd = etree.SubElement(ws, M + "sheetData")
    for row in sorted({r for r, _ in cells}):
        r_el = etree.SubElement(sd, M + "row", r=str(row))
        for (rr, cc) in sorted(k for k in cells if k[0] == row):
            kind, val, *style = cells[(rr, cc)]
            c = etree.SubElement(r_el, M + "c", r="%s%d" % (num_to_col(cc), rr))
            if style:
                c.set("s", str(style[0]))
            if kind == "s":
                c.set("t", "s")
            etree.SubElement(c, M + "v").text = str(val)
    if dv:
        dvs = etree.SubElement(ws, M + "dataValidations", count="1")
        d = etree.SubElement(dvs, M + "dataValidation", type="list", allowBlank="1",
                             showInputMessage="1", showErrorMessage="1", sqref=dv[0])
        etree.SubElement(d, M + "formula1").text = dv[1]
    etree.SubElement(ws, M + "pageMargins", left="0.7", right="0.7", top="0.75",
                     bottom="0.75", header="0.3", footer="0.3")
    tp = etree.SubElement(ws, M + "tableParts", count="1")
    etree.SubElement(tp, M + "tablePart", {R + "id": table_rid})
    return dump(ws)


def table_xml(table_id, name, ref, columns):
    t = etree.Element(M + "table", nsmap={None: NS_MAIN}, id=str(table_id), name=name,
                      displayName=name, ref=ref, totalsRowShown="0")
    etree.SubElement(t, M + "autoFilter", ref=ref)
    tc = etree.SubElement(t, M + "tableColumns", count=str(len(columns)))
    for i, col in enumerate(columns, start=1):
        etree.SubElement(tc, M + "tableColumn", id=str(i), name=col)
    etree.SubElement(t, M + "tableStyleInfo", name="TableStyleMedium2", showFirstColumn="0",
                     showLastColumn="0", showRowStripes="1", showColumnStripes="0")
    return dump(t)


def rels_xml(targets):
    root = etree.Element("{%s}Relationships" % NS_PKG_REL, nsmap={None: NS_PKG_REL})
    for i, (rtype, target) in enumerate(targets, start=1):
        etree.SubElement(root, "{%s}Relationship" % NS_PKG_REL, Id="rId%d" % i,
                         Type=rtype, Target=target)
    return dump(root)


def make_settings(sst, title_style, note_style):
    cells = {
        (1, 1): ("s", sst.add("Dashboard settings"), title_style),
        (2, 1): ("s", sst.add(
            "Manpower charge-to: a manpower line (Rate Category + IO Code) listed here "
            "appears as a locked row in the Charge To block instead of MP. "
            "Lines not listed are charged to MP."), note_style),
    }
    for i, name in enumerate(CHARGE_COLUMNS, start=1):
        cells[(4, i)] = ("s", sst.add(name))
    cells[(7, 1)] = ("s", sst.add(
        "Rate Category: as in tbl_ManpowerRate (RA, AHS, AHS (OT/PT), Nurse; "
        "case is ignored). IO Code: the manpower IO exactly as in the manpower "
        "source, e.g. TF25HIP136MP. Charge To: MP, OM, EQ, CO or OTHER."), note_style)
    cells[(8, 1)] = ("s", sst.add(
        "Changes apply the next time someone clicks Load Project."), note_style)
    return sheet_xml(cells, [26, 14, 16, 18, 12], "rId1",
                     dv=("E5:E500", '"MP,OM,EQ,CO,OTHER"'), tab_color="FF0D6B5F")


def make_store(sst):
    cells = {}
    for i, name in enumerate(STORE_COLUMNS, start=1):
        cells[(1, i)] = ("s", sst.add(name))
    cells[(1, 13)] = ("s", sst.add("Grid layout (rewritten by Load Project; do not edit)"))
    return sheet_xml(cells, [24, 12, 14, 10, 8, 30, 30, 11, 12, 18, 18], "rId1")


def edit_manpower_sheet(data):
    root = parse(data)
    sd = root.find(M + "sheetData")
    for r in sd.findall(M + "row"):
        if int(r.get("r")) == 15:
            cells = r.findall(M + "c")
            refs = [c.get("r") for c in cells]
            f = cells[0].find(M + "f") if cells else None
            if refs != ["P15"] or f is None or not f.text.startswith("SUM(qManpower_SelectedTest"):
                raise SystemExit("Unexpected content in PQManpower_SelectedTest row 15: %s" % refs)
            sd.remove(r)
    root.find(M + "dimension").set("ref", "A1:R13")
    return dump(root)


def rename_manpower_table(data):
    """Exact text replace, so the rest of this Power Query table part stays byte-identical."""
    old, new = MANPOWER_RENAME["table"]
    text = data.decode("utf-8")
    attrs = 'name="%s" displayName="%s"'
    if text.count(attrs % (old, old)) != 1:
        raise SystemExit("table15.xml is not %s" % old)
    return text.replace(attrs % (old, old), attrs % (new, new)).encode("utf-8")


def edit_app_props(data):
    old, new = MANPOWER_RENAME["sheet"]
    tag = "<vt:lpstr>%s</vt:lpstr>" % old
    text = data.decode("utf-8")
    return text.replace(tag, "<vt:lpstr>%s</vt:lpstr>" % new).encode("utf-8")


def edit_workbook(data):
    root = parse(data)
    sheets = root.find(M + "sheets")
    old_sheet, new_sheet = MANPOWER_RENAME["sheet"]
    renamed = [s for s in sheets if s.get("name") == old_sheet]
    if len(renamed) != 1:
        raise SystemExit("Sheet %s not found" % old_sheet)
    renamed[0].set("name", new_sheet)
    next_id = max(int(s.get("sheetId")) for s in sheets) + 1
    for offset, (name, state) in enumerate(
            [("Legacy_Timeline", "hidden"), ("Dashboard_Settings", None),
             ("ProjectionStore", "veryHidden")]):
        attrs = {"name": name, "sheetId": str(next_id + offset),
                 R + "id": "rIdV9%d" % (offset + 1)}
        if state:
            attrs["state"] = state
        etree.SubElement(sheets, M + "sheet", attrs)

    dn = root.find(M + "definedNames")
    existing = {d.get("name") for d in dn}
    for name, ref in NEW_NAMES:
        if name in existing:
            raise SystemExit("Name already exists: " + name)
        el = etree.Element(M + "definedName", name=name)
        el.text = ref
        later = [d for d in dn if d.get("name").lower() > name.lower()]
        if later:
            later[0].addprevious(el)
        else:
            dn.append(el)

    view = root.find(M + "bookViews").find(M + "workbookView")
    view.set("firstSheet", "0")
    view.set("activeTab", "0")
    root.find(M + "calcPr").set("fullCalcOnLoad", "1")
    return dump(root)


def edit_workbook_rels(data):
    root = parse(data)
    for rel in list(root):
        if rel.get("Target") == "calcChain.xml":
            root.remove(rel)
    for i, target in enumerate(["worksheets/sheet18.xml", "worksheets/sheet19.xml",
                                "worksheets/sheet20.xml"], start=1):
        etree.SubElement(root, "{%s}Relationship" % NS_PKG_REL, Id="rIdV9%d" % i,
                         Type=REL_WS, Target=target)
    return dump(root)


def edit_content_types(data):
    root = parse(data)
    for ov in list(root):
        if ov.get("PartName") in ("/xl/calcChain.xml", "/xl/comments1.xml"):
            root.remove(ov)
    for part in (LEGACY, SETTINGS, STORE):
        etree.SubElement(root, "{%s}Override" % NS_CT, PartName="/" + part, ContentType=CT_WS)
    for part in (T_CHARGE, T_STORE):
        etree.SubElement(root, "{%s}Override" % NS_CT, PartName="/" + part, ContentType=CT_TABLE)
    return dump(root)


def cell_style(dashboard_xml, ref):
    sd = parse(dashboard_xml).find(M + "sheetData")
    _, c = find_cell(sd, ref)
    return c.get("s")


def build(src, dst):
    zin = zipfile.ZipFile(src)
    parts = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    order = [i.filename for i in zin.infolist()]
    original_dashboard = parts[DASH]
    sst = SharedStrings(parts["xl/sharedStrings.xml"])

    out = dict(parts)
    out[DASH] = edit_dashboard(original_dashboard, sst)
    out[LEGACY] = make_legacy(original_dashboard, sst)
    out[SETTINGS] = make_settings(sst, cell_style(original_dashboard, "B5"),
                                  cell_style(original_dashboard, "C12"))
    out[STORE] = make_store(sst)
    out[MANPOWER_SHEET] = edit_manpower_sheet(parts[MANPOWER_SHEET])
    out[T_MANPOWER] = rename_manpower_table(parts[T_MANPOWER])
    out["docProps/app.xml"] = edit_app_props(parts["docProps/app.xml"])
    out[T_CHARGE] = table_xml(25, "tblManpowerChargeTo", "A4:E5", CHARGE_COLUMNS)
    out[T_STORE] = table_xml(26, "tblProjectionStore", "A1:K2", STORE_COLUMNS)
    out["xl/worksheets/_rels/sheet1.xml.rels"] = None
    out["xl/worksheets/_rels/sheet18.xml.rels"] = rels_xml(
        [(REL_TABLE, "../tables/table1.xml"), (REL_TABLE, "../tables/table2.xml")])
    out["xl/worksheets/_rels/sheet19.xml.rels"] = rels_xml([(REL_TABLE, "../tables/table18.xml")])
    out["xl/worksheets/_rels/sheet20.xml.rels"] = rels_xml([(REL_TABLE, "../tables/table19.xml")])
    out["xl/calcChain.xml"] = None
    out["xl/comments1.xml"] = None
    out["xl/drawings/vmlDrawing1.vml"] = None
    out["xl/workbook.xml"] = edit_workbook(parts["xl/workbook.xml"])
    out["xl/_rels/workbook.xml.rels"] = edit_workbook_rels(parts["xl/_rels/workbook.xml.rels"])
    out["[Content_Types].xml"] = edit_content_types(parts["[Content_Types].xml"])
    out["xl/sharedStrings.xml"] = sst.bytes()

    names = [n for n in order if out.get(n) is not None]
    names += [n for n in out if n not in parts and out[n] is not None]
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, out[n])
    return names


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    if sys.argv[1] == sys.argv[2]:
        raise SystemExit("Refusing to overwrite the source workbook.")
    written = build(sys.argv[1], sys.argv[2])
    print("Wrote %s (%d parts)" % (sys.argv[2], len(written)))
