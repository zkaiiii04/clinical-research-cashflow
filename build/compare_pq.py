"""Acceptance test 13: Power Query untouched after a Load/Save cycle in Excel.

Compares two workbooks (e.g. v0.8 and a v0.9_working that has been through
Load Project / Save Projections and saved by Excel). Excel re-encodes the
package when it saves, so this compares meaning, not bytes:
  * the M code (Formulas/Section1.m inside the DataMashup) is identical;
  * the same connections exist, with the same command text;
  * every query table loads to the same table name on the same sheet;
  * the saved file's own links hold (verify_v09_working.check_links): hidden
    ExternalData_n names, query tables, tables and FillTargets agree.

Usage:
    python3 -I compare_pq.py <baseline.xlsm> <after.xlsm>
"""
import os
import re
import sys
import zipfile

# python3 -I leaves the script's folder off sys.path; add it for the sibling module.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from verify_v09_working import check_links, m_code  # noqa: E402

def connections(z):
    if "xl/connections.xml" not in z.namelist():
        return {}
    xml = z.read("xl/connections.xml").decode("utf-8")
    out = {}
    for m in re.finditer(r'<connection [^>]*name="([^"]+)"[^>]*>(.*?)</connection>', xml, re.S):
        cmd = re.search(r'command="([^"]*)"', m.group(2))
        out[m.group(1)] = cmd.group(1) if cmd else ""
    return out


def query_tables(z):
    """table name -> sheet name, for every table loaded by a query."""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid_target = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    rid_target.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    out = {}
    for name, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        sheet_part = "xl/" + rid_target[rid].lstrip("/").replace("xl/", "")
        rel_part = sheet_part.replace("worksheets/", "worksheets/_rels/") + ".rels"
        if rel_part not in z.namelist():
            continue
        for target in re.findall(r'Target="([^"]+)"', z.read(rel_part).decode("utf-8")):
            if "tables/" in target:
                part = "xl/tables/" + target.split("tables/")[-1]
                t = z.read(part).decode("utf-8")
                if 'tableType="queryTable"' in t:
                    tname = re.search(r' name="([^"]+)"', t).group(1)
                    out[tname] = name
    return out


def main(a, b):
    za, zb = zipfile.ZipFile(a), zipfile.ZipFile(b)
    problems = []
    try:
        if m_code(za) != m_code(zb):
            problems.append("M code differs")
    except KeyError as e:
        problems.append("Power Query package missing: %s" % e)
    ca, cb = connections(za), connections(zb)
    if ca != cb:
        problems.append("Connections differ: %s" % sorted(set(ca.items()) ^ set(cb.items())))
    qa, qb = query_tables(za), query_tables(zb)
    if qa != qb:
        problems.append("Query table destinations differ: %s" % sorted(set(qa.items()) ^ set(qb.items())))
    print("Connections: %d, query tables: %d" % (len(cb), len(qb)))
    problems += check_links(zb)
    for p in problems:
        print("PROBLEM:", p)
    print("RESULT:", "PASS" if not problems else "FAIL")
    return 0 if not problems else 1


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    sys.exit(main(sys.argv[1], sys.argv[2]))
