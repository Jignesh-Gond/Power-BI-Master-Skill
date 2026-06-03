"""
PowerBI Automation Suite — Unified CLI
Consolidated from 52 scripts into one file.
Usage: python powerbi_automation.py <command> [args]
"""

import zipfile, json, os, shutil, uuid, time, re
import argparse, textwrap
from xml.etree import ElementTree as ET

# ─────────────────────────────────────────────
# CONSTANTS — THEMES, LAYOUT, DISPLAY UNITS
# ─────────────────────────────────────────────
TITLE_BG = "#E6F3FF"
KPI_VALUE_COLOR = "#17614A"
BORDER_COLOR = "#CCCCCC"
CHART_TITLE_SIZE = "13D"
KPI_TITLE_SIZE = "14D"
KPI_VALUE_SIZE = "22D"
BORDER_RADIUS = "8D"
SLICER_MIN_HEIGHT = 60
CONTAINER_RADIUS = "8D"
DISPLAY_UNITS_NONE = "0D"
DISPLAY_UNITS_THOUSANDS = "1L"

L = lambda v: {"expr": {"Literal": {"Value": v}}}

BASE_MEASURES_DICT = {
    "Total Revenue": "CALCULATE(SUM('SalesTable'[Final Rate]) * SUM('SalesTable'[QTY])) / 10000000",
    "Total QTY": "SUM('SalesTable'[QTY])",
    "Total Products": "DISTINCTCOUNT('SalesTable'[Product])",
    "Customers Covered": "DISTINCTCOUNT('SalesTable'[Customer Name])",
    "Active Sales Reps": "DISTINCTCOUNT('SalesTable'[Sales Rep Name])",
    "Avg Final Rate": "DIVIDE(CALCULATE(SUM('SalesTable'[Final Rate]) * SUM('SalesTable'[QTY])), SUM('SalesTable'[QTY]))",
    "Avg PTR Rate": "DIVIDE(CALCULATE(SUM('SalesTable'[PTR Rate]) * SUM('SalesTable'[QTY])), SUM('SalesTable'[QTY]))",
    "Total Revenue (CR)": "CALCULATE(SUM('SalesTable'[Final Rate]) * SUM('SalesTable'[QTY])) / 10000000",
    "Revenue (Lakh Numeric)": "CALCULATE(SUM('SalesTable'[Final Rate]) * SUM('SalesTable'[QTY])) / 100000",
    "QTY (Lakh Numeric)": "SUM('SalesTable'[QTY]) / 100000",
    "Avg Final Rate Numeric": "DIVIDE(CALCULATE(SUM('SalesTable'[Final Rate]) * SUM('SalesTable'[QTY])), SUM('SalesTable'[QTY]))",
    "Avg PTR Rate Numeric": "DIVIDE(CALCULATE(SUM('SalesTable'[PTR Rate]) * SUM('SalesTable'[QTY])), SUM('SalesTable'[QTY]))",
    "Total Revenue Indian": "VAR Val = CALCULATE(SUM('SalesTable'[Final Rate]) * SUM('SalesTable'[QTY])) / 10000000 VAR SafeLog = IF(Val > 0, FLOOR(LOG10(Val + 0.0000001), 1), 0) VAR GroupsAfter3 = MIN(3, MAX(0, ROUNDUP((SafeLog - 2) / 2, 0))) VAR Pattern = IF(GroupsAfter3 = 0, \"#,##0\", REPT(\"#,##\\\\,\", GroupsAfter3) & \"##0\") RETURN IF(Val = 0 || ISBLANK(Val), \"0\", FORMAT(Val, Pattern))",
    "Total QTY Indian": "VAR Val = SUM('SalesTable'[QTY]) VAR SafeLog = IF(Val > 0, FLOOR(LOG10(Val + 0.0000001), 1), 0) VAR GroupsAfter3 = MIN(3, MAX(0, ROUNDUP((SafeLog - 2) / 2, 0))) VAR Pattern = IF(GroupsAfter3 = 0, \"#,##0\", REPT(\"#,##\\\\,\", GroupsAfter3) & \"##0\") RETURN IF(Val = 0 || ISBLANK(Val), \"0\", FORMAT(Val, Pattern))",
    "Customers Covered (Display)": "FORMAT([Customers Covered], \"#,##0\")",
    "Total Products (Display)": "FORMAT([Total Products], \"#,##0\")",
    "Active Sales Reps (Display)": "FORMAT([Active Sales Reps], \"#,##0\")",
    "Total QTY (Display)": "FORMAT([Total QTY], \"#,##0\")",
    "Avg Final Rate (Display)": "FORMAT([Avg Final Rate], \"#,##0\")",
    "Avg PTR Rate (Display)": "FORMAT([Avg PTR Rate], \"#,##0\")",
    "Revenue (Cr Numeric)": "CALCULATE(SUM('SalesTable'[Final Rate]) * SUM('SalesTable'[QTY])) / 10000000",
    "Customer Sales %": "VAR AllRev = CALCULATE(SUM('SalesTable'[Final Rate]) * SUM('SalesTable'[QTY]), ALLEXCEPT('SalesTable', 'SalesTable'[Sales Rep Name])) VAR CustRev = CALCULATE(SUM('SalesTable'[Final Rate]) * SUM('SalesTable'[QTY]), ALLEXCEPT('SalesTable', 'SalesTable'[Sales Rep Name], 'SalesTable'[Customer Name])) RETURN DIVIDE(CustRev, AllRev)",
    "Partner Sales %": "VAR PartnerRev = CALCULATE(SUM('SalesTable'[Final Rate]) * SUM('SalesTable'[QTY]), ALLEXCEPT('SalesTable', 'SalesTable'[Sales Rep Name], 'SalesTable'[Partner Name])) VAR AllRev = CALCULATE(SUM('SalesTable'[Final Rate]) * SUM('SalesTable'[QTY]), ALLEXCEPT('SalesTable', 'SalesTable'[Sales Rep Name])) RETURN DIVIDE(PartnerRev, AllRev)",
    "Sales Performance Detail": "VAR CustSales = [Customer Sales %] VAR PartnerSales = [Partner Sales %] RETURN CustSales & \" | \" & PartnerSales",
    "Dynamic Growth %": "VAR Prev = CALCULATE([Total Revenue], DATEADD('Date'[Date], -1, MONTH)) RETURN DIVIDE([Total Revenue] - Prev, Prev)",
    "Growth Category": "SWITCH(TRUE(), [Dynamic Growth %] > 0.05, \"High Growth\", [Dynamic Growth %] > 0, \"Growth\", [Dynamic Growth %] > -0.05, \"Stable\", \"Decline\")",
    "Growth Arrow": "SWITCH(TRUE(), [Dynamic Growth %] > 0.05, \"↑↑\", [Dynamic Growth %] > 0, \"↑\", [Dynamic Growth %] > -0.05, \"→\", \"↓\")",
    "Growth Color": "SWITCH(TRUE(), [Dynamic Growth %] > 0.05, \"DarkGreen\", [Dynamic Growth %] > 0, \"Green\", [Dynamic Growth %] > -0.05, \"Orange\", \"Red\")",
    "Region Revenue %": "VAR RegionRev = CALCULATE(SUM('SalesTable'[Final Rate]) * SUM('SalesTable'[QTY]), ALLEXCEPT('SalesTable', 'SalesTable'[Region])) VAR TotalRev = CALCULATE(SUM('SalesTable'[Final Rate]) * SUM('SalesTable'[QTY]), ALL('SalesTable')) RETURN DIVIDE(RegionRev, TotalRev)",
    "Dual KPI Display": "VAR MainVal = [Total Revenue (CR)] VAR Growth = [Dynamic Growth %] RETURN FORMAT(MainVal, \"#,##0.00\") & \" (\" & FORMAT(Growth, \"+0.0%;-0.0%\") & \")\"",
    "Dual KPI Color": "IF([Dynamic Growth %] >= 0, \"#17614A\", \"#DC3545\")",
}

# ─────────────────────────────────────────────
# CORE UTILITIES
# ─────────────────────────────────────────────
def extract_pbix(src, tmp_dir):
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    with zipfile.ZipFile(src, 'r') as z:
        z.extractall(tmp_dir)

def repack_pbix(tmp_dir, output, data_model_stored=True):
    if os.path.exists(output):
        os.remove(output)
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(tmp_dir):
            for fn in files:
                fp = os.path.join(root, fn)
                arc = os.path.relpath(fp, tmp_dir)
                if fn == "DataModel" and data_model_stored:
                    zout.write(fp, arc, compress_type=zipfile.ZIP_STORED)
                else:
                    zout.write(fp, arc, compress_type=zipfile.ZIP_DEFLATED)

def pbix_entries(path):
    with zipfile.ZipFile(path, 'r') as z:
        return [(n, z.getinfo(n).file_size) for n in z.namelist()]

def remove_security_bindings(tmp_dir):
    for name in ["SecurityBindings", "SecurityBindings.xml"]:
        p = os.path.join(tmp_dir, name)
        if os.path.isfile(p):
            os.remove(p)
        elif os.path.isdir(p):
            shutil.rmtree(p)
    ct_path = os.path.join(tmp_dir, "[Content_Types].xml")
    with open(ct_path, 'r', encoding='utf-8') as f:
        ct = f.read()
    ct = ct.replace('<Override PartName="/SecurityBindings/SecurityBindings" ContentType="application/octet-stream"/>', "")
    ct = ct.replace('<Override PartName="/SecurityBindings/SecurityBindings.xml" ContentType="application/xml"/>', "")
    with open(ct_path, 'w', encoding='utf-8') as f:
        f.write(ct)
    return ct

def load_pages_json(tmp_dir):
    pf = os.path.join(tmp_dir, 'Report/definition/pages/pages.json')
    with open(pf, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_pages_json(tmp_dir, data):
    pf = os.path.join(tmp_dir, 'Report/definition/pages/pages.json')
    with open(pf, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)

def get_page_name(tmp_dir, pid):
    pf = os.path.join(tmp_dir, f'Report/definition/pages/{pid}/page.json')
    if not os.path.exists(pf):
        return None
    with open(pf, 'r', encoding='utf-8') as f:
        return json.load(f).get('displayName')

def iter_visuals(tmp_dir, pid):
    vis_dir = os.path.join(tmp_dir, f'Report/definition/pages/{pid}/visuals')
    if not os.path.exists(vis_dir):
        return
    for vid in os.listdir(vis_dir):
        vf = os.path.join(vis_dir, vid, 'visual.json')
        if os.path.exists(vf):
            with open(vf, 'r', encoding='utf-8') as f:
                yield vid, json.load(f), vf

def write_visual(vf, data):
    with open(vf, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)

# TOM Helpers (SSAS)
def find_port():
    import subprocess, re
    r = subprocess.run(['netstat', '-ano'], capture_output=True, text=True)
    ports = re.findall(r'MSAS\s+(\d+)', r.stdout, re.IGNORECASE)
    if not ports:
        ports = re.findall(r':(\d+)\s+\d+\.\d+\.\d+\.\d+:\d+\s+LISTENING\s+\d+', r.stdout)
        pbi_pids = set()
        try:
            t = subprocess.run(['tasklist', '/fi', 'imagename eq PowerBIDesktop.exe', '/fo', 'csv'],
                               capture_output=True, text=True)
            for line in t.stdout.strip().split('\n')[1:]:
                parts = line.split(',')
                if len(parts) >= 2:
                    pbi_pids.add(parts[1].strip('"'))
        except:
            pass
        port_pids = {}
        for line in r.stdout.split('\n'):
            m = re.search(r':(\d+)\s+\S+\s+LISTENING\s+(\d+)', line)
            if m:
                port_pids[m.group(1)] = m.group(2)
        for p, pid in port_pids.items():
            if pid in pbi_pids:
                ports.append(p)
    return ports[0] if ports else None

def connect_tom(port):
    from Microsoft.AnalysisServices.Tabular import Server
    s = Server()
    s.Connect(f"Data Source=localhost:{port}")
    return s

def find_port_retry(retries=60, delay=2):
    for i in range(retries):
        p = find_port()
        if p:
            return p
        time.sleep(delay)
    raise RuntimeError(f"SSAS port not found after {retries} attempts")

def send_ctrl_s(window_title=None):
    try:
        import win32gui, win32con, ctypes
        if window_title:
            hwnd = win32gui.FindWindow(None, window_title)
        else:
            hwnd = win32gui.FindWindow("Power B I Desktop", None)
            if not hwnd:
                for h in range(10000):
                    if win32gui.IsWindow(h) and "Power BI" in win32gui.GetWindowText(h):
                        hwnd = h; break
        if hwnd:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.5)
            ctypes.windll.user32.keybd_event(0x11, 0, 0, 0)
            ctypes.windll.user32.keybd_event(0x53, 0, 0, 0)
            time.sleep(0.1)
            ctypes.windll.user32.keybd_event(0x53, 0, 2, 0)
            ctypes.windll.user32.keybd_event(0x11, 0, 2, 0)
            return True
    except Exception as e:
        print(f"Ctrl+S failed: {e}")
    return False

# ─────────────────────────────────────────────
# VISUAL HELPER BUILDERS
# ─────────────────────────────────────────────
def col_field(ent, prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": ent}}, "Property": prop}}

def agg_field(ent, prop, func=1):
    return {"Aggregation": {"Expression": col_field(ent, prop), "Function": func}}

def meas_field(ent, prop):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": ent}}, "Property": prop}}

def make_card(vid, measure_table, measure_name, title, x, y, w, h, z):
    return {
        "name": vid,
        "position": {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z},
        "visual": {
            "visualType": "card",
            "query": {"queryRegion": {"Values": {"projections": [
                {"field": meas_field(measure_table, measure_name), "queryRef": f"{measure_table}.{measure_name}", "nativeQueryRef": measure_name}
            ]}}},
            "objects": {
                "categoryLabels": [{"properties": {"show": L("false")}}],
                "labels": [{"properties": {
                    "fontSize": L(KPI_VALUE_SIZE),
                    "color": {"solid": {"color": L(f"'{KPI_VALUE_COLOR}'")}},
                    "labelDisplayUnits": L("1L"),
                    "labelPrecision": L("0L"),
                    "horizontalAlignment": L("'center'")
                }}],
                "wordWrap": [{"properties": {"show": L("true")}}],
                "subTitle": [{"properties": {"show": L("false"), "text": L("''"), "fontSize": L("1D")}}],
                "general": [{"properties": {"autoSubtitle": L("false")}}]
            },
            "visualContainerObjects": {
                "title": [{"properties": {
                    "text": L(f"'{title}'"), "show": L("true"),
                    "titleWrap": L("true"),
                    "fontColor": {"solid": {"color": L("'#333333'")}},
                    "background": {"solid": {"color": L(f"'{TITLE_BG}'")}},
                    "alignment": L("'center'"), "fontSize": L(KPI_TITLE_SIZE)
                }}],
                "visualHeader": [{"properties": {"show": L("true"), "transparency": L("100D")}}],
                "border": [{"properties": {"show": L("true"), "radius": L("8D"), "width": L("1D"), "color": {"solid": {"color": L("'#17614A'")}}}}],
                "dropShadow": [{"properties": {"show": L("true")}}],
                "padding": [{"properties": {"top": L("0D"), "bottom": L("0D"), "left": L("0D"), "right": L("0D")}}],
                "subTitle": [{"properties": {"show": L("false"), "text": L("''")}}]
            },
            "drillFilterOtherVisuals": True
        }
    }

def make_slicer(vid, title, x, y, w, h, z, ent, prop):
    return {
        "name": vid,
        "position": {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z},
        "visual": {
            "visualType": "slicer",
            "query": {"queryRegion": {"Values": {"projections": [
                {"field": col_field(ent, prop), "queryRef": prop}
            ]}}},
            "objects": {
                "general": [{"properties": {"selfFilterEnabled": L("true"), "outlineColor": {"solid": {"color": L("'#6E929A'")}}, "orientation": L("1D")}}],
                "selection": [{"properties": {"selectAllCheckboxEnabled": L("true"), "singleSelect": L("false")}}],
                "data": [{"properties": {"mode": L("'Dropdown'")}}],
                "outline": [{"properties": {"show": L("true")}}],
                "header": [{"properties": {"text": L(f"'{title}'"), "textSize": L("11D"), "bold": L("false"), "fontColor": {"solid": {"color": L("'#333333'")}}}}],
                "items": [{"properties": {"textSize": L("9D"), "fontColor": {"solid": {"color": L("'#333333'")}}, "background": {"solid": {"color": L("'#E6F3FF'")}}}}]
            },
            "visualContainerObjects": {
                "title": [{"properties": {"show": L("false")}}],
                "visualHeader": [{"properties": {"show": L("true"), "transparency": L("100D")}}],
                "border": [{"properties": {"show": L("true"), "radius": L("6D"), "width": L("1D"), "color": {"solid": {"color": L(f"'{BORDER_COLOR}'")}}}}],
                "divider": [{"properties": {"show": L("false")}}],
                "spacing": [{"properties": {"verticalSpacing": L("3D")}}],
                "padding": [{"properties": {"top": L("4D")}}]
            },
            "drillFilterOtherVisuals": True
        }
    }

def make_chart(vid, vtype, title, x, y, w, h, z, cat_ent, cat_prop, val_table, val_prop, orient=1):
    obj = {
        "name": vid,
        "position": {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z},
        "visual": {
            "visualType": vtype,
            "query": {"queryRegion": {
                "Category": {"projections": [{
                    "field": agg_field(cat_ent, cat_prop),
                    "queryRef": f"{cat_ent}.{cat_prop}",
                    "nativeQueryRef": cat_prop, "active": True
                }]},
                "Y": {"projections": [{
                    "field": meas_field(val_table, val_prop),
                    "queryRef": "Val0", "nativeQueryRef": val_prop
                }]}
            }},
            "objects": {
                "legend": [{"properties": {"show": L("true"), "position": L("'BottomAuto'")}}],
                "labels": [{"properties": {"show": L("true"), "labelPosition": L("'OutsideEnd'"), "fontSize": L("10D"), "color": {"solid": {"color": L("'#333333'")}}}}],
                "categoryAxis": [{"properties": {"labelColor": {"solid": {"color": L("'#333333'")}}, "showAxisTitle": L("false"), "fontSize": L("10D"), "bold": L("false"), "concatenateLabels": L("true")}}],
                "valueAxis": [{"properties": {"labelColor": {"solid": {"color": L("'#333333'")}}, "showAxisTitle": L("false"), "show": L("false"), "gridlineShow": L("false")}}]
            },
            "visualContainerObjects": {
                "title": [{"properties": {"text": L(f"'{title}'"), "show": L("true"), "titleWrap": L("true"), "fontColor": {"solid": {"color": L("'#333333'")}}, "background": {"solid": {"color": L(f"'{TITLE_BG}'")}}, "alignment": L("'center'"), "fontSize": L(CHART_TITLE_SIZE)}}],
                "visualHeader": [{"properties": {"show": L("true"), "transparency": L("100D")}}],
                "border": [{"properties": {"show": L("true"), "radius": L(BORDER_RADIUS), "width": L("1D"), "color": {"solid": {"color": L(f"'{BORDER_COLOR}'")}}}}],
                "dropShadow": [{"properties": {"show": L("true")}}]
            },
            "drillFilterOtherVisuals": True
        }
    }
    if vtype in ("barChart", "columnChart"):
        obj["visual"]["objects"]["general"] = [{"properties": {"orientation": L(f"{orient}D")}}]
    return obj

def create_ranking_page(name, displayName, pid, columns, measures, sort_measure):
    page = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/1.0.0/schema.json",
        "name": pid, "displayName": displayName, "displayOption": "FitToPage", "height": 720, "width": 1280
    }
    vid = "rank_table_" + name
    v = {
        "name": vid,
        "position": {"x": 10, "y": 10, "z": 0, "height": 700, "width": 1260, "tabOrder": 0},
        "visual": {
            "visualType": "tableEx",
            "query": {"queryRegion": {"Values": {"projections": []}}},
            "objects": {"values": [{"properties": {"fontSize": L("11D")}}]},
            "visualContainerObjects": {
                "title": [{"properties": {
                    "text": L(f"'{displayName}'"), "show": L("true"),
                    "background": {"solid": {"color": L(f"'{TITLE_BG}'")}},
                    "alignment": L("'center'"), "fontSize": L("16D")
                }}],
                "subtitle": [{"properties": {"show": L("false")}}],
                "border": [{"properties": {"show": L("true"), "radius": L(BORDER_RADIUS), "color": {"solid": {"color": L(f"'{KPI_VALUE_COLOR}'")}}}}]
            }
        },
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.8.0/schema.json"
    }
    projs = v['visual']['query']['queryRegion']['Values']['projections']
    for c_ent, c_prop in columns:
        projs.append({"field": col_field(c_ent, c_prop), "queryRef": c_prop})
    for m_ent, m_prop in measures:
        projs.append({"field": meas_field(m_ent, m_prop), "queryRef": m_prop})
    s_ent, s_prop = sort_measure
    projs.append({"field": meas_field(s_ent, s_prop), "queryRef": "SortCol", "isHidden": True})
    v['visual']['query']['queryRegion']['Values']['orderBy'] = [
        {"direction": 2, "field": meas_field(s_ent, s_prop)}
    ]
    return page, v

# ─────────────────────────────────────────────
# COMMAND: analyze
# ─────────────────────────────────────────────
def cmd_analyze(args):
    """Show full structure of a PBIX file: pages, visuals, fields."""
    print(f"\n{'='*60}")
    print(f"  Analyzing: {args.pbix}")
    print(f"{'='*60}")
    with zipfile.ZipFile(args.pbix, 'r') as z:
        names = z.namelist()
        page_files = [n for n in names if n.endswith('page.json') and 'pages/' in n]
        visual_files = [n for n in names if n.endswith('visual.json')]
        print(f"\n  Total entries: {len(names)}")
        print(f"  Pages: {len(page_files)}, Visuals: {len(visual_files)}")
        for n in sorted(names):
            info = z.getinfo(n)
            print(f"    {n} ({info.file_size} bytes)")
        if 'DataModelSchema' in names:
            schema = json.loads(z.read('DataModelSchema'))
            measures = schema.get('model', {}).get('measures', [])
            print(f"\n  DAX Measures: {len(measures)}")
            for m in measures:
                print(f"    [{m.get('table')}] {m.get('name')} = {m.get('expression', '')[:100]}")
        if 'Report/definition/pages/pages.json' in names:
            pages = json.loads(z.read('Report/definition/pages/pages.json'))
            print(f"\n  Page Order ({len(pages.get('pageOrder', []))}):")
            for pid in pages.get('pageOrder', []):
                pg = json.loads(z.read(f'Report/definition/pages/{pid}/page.json'))
                print(f"    {pg.get('displayName')}: {pg.get('width')}x{pg.get('height')}")

def cmd_analyze_visuals(args):
    """Read visual.json files and show type/title/fields."""
    tmp = "_analyze_vis_tmp"
    extract_pbix(args.pbix, tmp)
    pages = load_pages_json(tmp)
    for pid in pages['pageOrder']:
        pname = get_page_name(tmp, pid)
        print(f"\n--- Page: {pname} ({pid}) ---")
        for vid, v_json, vf in iter_visuals(tmp, pid):
            v = v_json.get('visual', {})
            vtype = v.get('visualType', '?')
            vco = v.get('visualContainerObjects', {})
            title = '?'
            if 'title' in vco:
                t = vco['title'][0].get('properties', {}).get('text', {})
                title = t.get('expr', {}).get('Literal', {}).get('Value', '?')
            print(f"  [{vid}] Type={vtype} Title={title}")
            qs = v.get('query', {}).get('queryRegion', {})
            for shelf_name, shelf in qs.items():
                for p in shelf.get('projections', []):
                    f = p.get('field', {})
                    for ftype, finfo in f.items():
                        print(f"    {shelf_name}: {ftype} {finfo.get('Property', '')}")
    shutil.rmtree(tmp)

# ─────────────────────────────────────────────
# COMMAND: verify
# ─────────────────────────────────────────────
def cmd_verify(args):
    """Verify PBIX integrity: pages, positions, overflow, DataModel."""
    tmp = "_verify_tmp"
    extract_pbix(args.pbix, tmp)
    pages = load_pages_json(tmp)
    overflow = False
    for pid in pages['pageOrder']:
        pname = get_page_name(tmp, pid)
        if not pname: continue
        pp = os.path.join(tmp, f'Report/definition/pages/{pid}/page.json')
        with open(pp) as f:
            pg = json.load(f)
        pw, ph = pg.get('width', 1280), pg.get('height', 720)
        vis_count = 0
        for vid, v_json, vf in iter_visuals(tmp, pid):
            vis_count += 1
            p = v_json.get('position', {})
            rx = p.get('x', 0) + p.get('width', 0)
            ry = p.get('y', 0) + p.get('height', 0)
            if rx > pw + 5 and p.get('x', 0) < pw:  # only flag if started within page AND overflowed
                print(f"  OVERFLOW: {pname}/{vid} ends at ({rx},{ry}) > ({pw},{ph})")
                overflow = True
        dm_size = 0
        for root, dirs, files in os.walk(tmp):
            for fn in files:
                if fn == "DataModel":
                    dm_size = os.path.getsize(os.path.join(root, fn))
        print(f"  {pname}: {pw}x{ph} ({vis_count} visuals) DM={dm_size} bytes")
    if not overflow:
        print("  No overflow detected.")
    shutil.rmtree(tmp)

# ─────────────────────────────────────────────
# COMMAND: pages (list/reorder/remove)
# ─────────────────────────────────────────────
def cmd_list_pages(args):
    """List all pages in a PBIX file."""
    tmp = "_list_pages_tmp"
    extract_pbix(args.pbix, tmp)
    pages = load_pages_json(tmp)
    for pid in pages['pageOrder']:
        pname = get_page_name(tmp, pid)
        print(f"  {pname} ({pid})")
    shutil.rmtree(tmp)

def cmd_reorder_pages(args):
    """Reorder pages. Provide comma-separated display names."""
    tmp = "_reorder_tmp"
    extract_pbix(args.pbix, tmp)
    pages = load_pages_json(tmp)
    name_order = [n.strip() for n in args.order.split(',')]
    pid_map = {}
    for pid in pages['pageOrder']:
        pname = get_page_name(tmp, pid)
        if pname: pid_map[pname] = pid
    new_order = []
    for n in name_order:
        if n in pid_map:
            new_order.append(pid_map[n])
    for pid in pages['pageOrder']:
        if pid not in new_order:
            new_order.append(pid)
    pages['pageOrder'] = new_order
    save_pages_json(tmp, pages)
    remove_security_bindings(tmp)
    repack_pbix(tmp, args.output or args.pbix)
    print(f"Reordered pages: {[get_page_name(tmp, p) for p in new_order]}")

def cmd_remove_pages(args):
    """Remove pages matching given display names (comma-separated)."""
    tmp = "_remove_pages_tmp"
    extract_pbix(args.pbix, tmp)
    pages = load_pages_json(tmp)
    names_to_remove = [n.strip() for n in args.names.split(',')]
    keep = []
    removed = []
    for pid in pages['pageOrder']:
        pname = get_page_name(tmp, pid)
        if pname and pname in names_to_remove:
            removed.append(pname)
            pdir = os.path.join(tmp, f'Report/definition/pages/{pid}')
            if os.path.exists(pdir):
                shutil.rmtree(pdir)
        else:
            keep.append(pid)
    pages['pageOrder'] = keep
    save_pages_json(tmp, pages)
    remove_security_bindings(tmp)
    repack_pbix(tmp, args.output or args.pbix)
    print(f"Removed pages: {removed}")

# ─────────────────────────────────────────────
# COMMAND: add-page
# ─────────────────────────────────────────────
def cmd_add_page(args):
    """Add a new blank page to the PBIX."""
    tmp = "_add_page_tmp"
    extract_pbix(args.pbix, tmp)
    pages = load_pages_json(tmp)
    new_pid = uuid.uuid4().hex[:20]
    pdir = os.path.join(tmp, f'Report/definition/pages/{new_pid}')
    os.makedirs(os.path.join(pdir, 'visuals'), exist_ok=True)
    with open(os.path.join(pdir, 'page.json'), 'w', encoding='utf-8') as f:
        json.dump({
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/1.0.0/schema.json",
            "name": new_pid, "displayName": args.name,
            "displayOption": "FitToPage", "height": 720, "width": 1280
        }, f, indent=2)
    pages['pageOrder'].append(new_pid)
    save_pages_json(tmp, pages)
    ct_path = os.path.join(tmp, "[Content_Types].xml")
    with open(ct_path, 'r', encoding='utf-8') as f:
        ct = f.read()
    override = f'<Override PartName="/Report/definition/pages/{new_pid}/page.json" ContentType="application/json"/>'
    if override not in ct:
        ct = ct.replace("</Types>", override + "\n</Types>")
    with open(ct_path, 'w', encoding='utf-8') as f:
        f.write(ct)
    remove_security_bindings(tmp)
    repack_pbix(tmp, args.output or args.pbix)
    print(f"Added page: {args.name} ({new_pid})")

# ─────────────────────────────────────────────
# COMMAND: fix-visuals (unified visual fixer)
# ─────────────────────────────────────────────
def fix_visual(v_json, disable_subtitle=True, fix_kpi=True, display_units=None, fix_aggregations=False):
    v = v_json.get('visual', {})
    vtype = v.get('visualType')
    vco = v.setdefault('visualContainerObjects', {})
    objs = v.setdefault('objects', {})

    if disable_subtitle:
        vco['subTitle'] = [{"properties": {"show": L("false"), "text": L("''")}}]
        objs.setdefault('subTitle', [{}])[0].setdefault('properties', {})['show'] = L("false")
        objs.setdefault('general', [{}])[0].setdefault('properties', {})['autoSubtitle'] = L("false")

    if fix_kpi and vtype == 'card':
        title = vco.setdefault('title', [{}])[0]
        t_props = title.setdefault('properties', {})
        t_props['show'] = L("true")
        t_props['alignment'] = L("'center'")
        t_props['fontSize'] = L(KPI_TITLE_SIZE)
        t_props['background'] = {"solid": {"color": L(f"'{TITLE_BG}'")}}
        padding = vco.setdefault('padding', [{}])[0]
        p_props = padding.setdefault('properties', {})
        p_props['top'] = L("0D")
        p_props['bottom'] = L("0D")
        p_props['left'] = L("0D")
        p_props['right'] = L("0D")
        cat = objs.setdefault('categoryLabels', [{}])[0]
        cat.setdefault('properties', {})['show'] = L("false")
        labels = objs.setdefault('labels', [{}])[0]
        l_props = labels.setdefault('properties', {})
        l_props['fontSize'] = L(KPI_VALUE_SIZE)
        l_props['horizontalAlignment'] = L("'center'")
        l_props['color'] = {"solid": {"color": L(f"'{KPI_VALUE_COLOR}'")}}

    if display_units:
        for target in ['valueAxis', 'categoryAxis', 'labels']:
            if target in objs:
                for item in objs[target]:
                    props = item.setdefault('properties', {})
                    if 'labelDisplayUnits' in props or target == 'labels':
                        props['labelDisplayUnits'] = L(display_units)
                    if 'displayUnits' in props:
                        props['displayUnits'] = L(display_units)

    if fix_aggregations:
        MEASURE_MAP = {
            'Customer Name': 'Total QTY',
            'Sales Rep Name': 'Total Revenue',
            'Region': 'Total QTY',
            'Business Unit': 'Total QTY',
            'Product': 'Total Revenue (CR)',
            'Segment': 'Total QTY',
        }
        MEASURE_TABLE = 'Measure'
        qs = v.get('query', {}).get('queryRegion', {})
        for shelf in qs.values():
            for p in shelf.get('projections', []):
                f = p.get('field', {})
                if 'Aggregation' in f:
                    col_name = f['Aggregation']['Expression']['Property']
                    if col_name in MEASURE_MAP:
                        p['field'] = meas_field(MEASURE_TABLE, MEASURE_MAP[col_name])
                        p['queryRef'] = MEASURE_MAP[col_name]

    return v_json

def cmd_fix_visuals(args):
    """Fix all visuals: disable subtitles, fix KPI cards, set display units."""
    tmp = f"_fix_visuals_tmp"
    extract_pbix(args.pbix, tmp)
    pages = load_pages_json(tmp)
    display_units = args.display_units or DISPLAY_UNITS_NONE

    for pid in pages['pageOrder']:
        for vid, v_json, vf in iter_visuals(tmp, pid):
            v_json = fix_visual(v_json, disable_subtitle=True, fix_kpi=True,
                                display_units=display_units, fix_aggregations=args.fix_aggs)
            write_visual(vf, v_json)
    remove_security_bindings(tmp)
    repack_pbix(tmp, args.output or args.pbix)
    print(f"Fixed visuals in all pages. Output: {args.output or args.pbix}")

# ─────────────────────────────────────────────
# COMMAND: align-visuals
# ─────────────────────────────────────────────
def cmd_align_visuals(args):
    """Align visuals in a grid layout (6 KPI, 5 slicer, 4+4+3 charts)."""
    tmp = "_align_visuals_tmp"
    extract_pbix(args.pbix, tmp)
    pages = load_pages_json(tmp)

    card_positions = {
        "v001": (30, 5, 195, 75, 0), "v002": (235, 5, 195, 75, 1),
        "v003": (440, 5, 195, 75, 2), "v004": (645, 5, 195, 75, 3),
        "v005": (850, 5, 195, 75, 4), "v006": (1055, 5, 195, 75, 5),
    }
    slicer_positions = {
        "v007": (30, 90, 232, 40, 6), "v008": (277, 90, 232, 40, 7),
        "v009": (524, 90, 232, 40, 8), "v010": (771, 90, 232, 40, 9),
        "v011": (1018, 90, 232, 40, 10),
    }
    row3_positions = {
        "v012": (25, 145, 300, 175, 11), "v013": (335, 145, 300, 175, 12),
        "v014": (645, 145, 300, 175, 13), "v017": (955, 145, 300, 175, 14),
    }
    row4_positions = {
        "v016": (25, 335, 300, 175, 15), "v015": (335, 335, 300, 175, 16),
        "v019": (645, 335, 300, 175, 17), "v020": (955, 335, 300, 175, 18),
    }
    row5_positions = {
        "v018": (20, 525, 400, 185, 19), "v022": (430, 525, 400, 185, 20),
        "v021": (840, 525, 400, 185, 21),
    }
    all_positions = {}
    all_positions.update(card_positions)
    all_positions.update(slicer_positions)
    all_positions.update(row3_positions)
    all_positions.update(row4_positions)
    all_positions.update(row5_positions)

    for pid in pages['pageOrder']:
        for vid, v_json, vf in iter_visuals(tmp, pid):
            if vid in all_positions:
                x, y, w, h, z = all_positions[vid]
                v_json['position'] = {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z}
                write_visual(vf, v_json)
    remove_security_bindings(tmp)
    repack_pbix(tmp, args.output or args.pbix)
    print(f"Aligned visuals. Output: {args.output or args.pbix}")

# ─────────────────────────────────────────────
# COMMAND: add-ranking-page
# ─────────────────────────────────────────────
def cmd_add_ranking_page(args):
    """Add a Customer/Sales Rep ranking page with table visual."""
    tmp = "_ranking_tmp"
    extract_pbix(args.pbix, tmp)
    pages = load_pages_json(tmp)
    ct_path = os.path.join(tmp, "[Content_Types].xml")
    with open(ct_path, 'r', encoding='utf-8') as f:
        ct = f.read()

    ranking_configs = [
        ("Customer", "Customer Performance Ranking", uuid.uuid4().hex[:20],
         [("Customer Name (Combine Multiple)", "Customer Name")],
         [("Measure", "Total QTY Indian"), ("Measure", "Total Revenue Indian")],
         ("Measure", "Total Revenue")),
        ("MR", "Sales Performance & Splits", uuid.uuid4().hex[:20],
         [("SalesTable", "Sales Rep Name")],
         [("Measure", "Total Revenue (CR)"), ("Measure", "Customer Sales %"),
          ("Measure", "Partner Sales %"), ("Measure", "Sales Performance Detail")],
         ("Measure", "Total Revenue")),
    ]

    overrides = []
    for rname, rdisp, rpid, rcols, rmeas, rsort in ranking_configs:
        page_json, vis_json = create_ranking_page(rname, rdisp, rpid, rcols, rmeas, rsort)
        pdir = os.path.join(tmp, f"Report/definition/pages/{rpid}")
        vdir = os.path.join(pdir, "visuals", vis_json['name'])
        os.makedirs(vdir, exist_ok=True)
        with open(os.path.join(pdir, "page.json"), 'w', encoding='utf-8') as f:
            json.dump(page_json, f, indent=2)
        with open(os.path.join(vdir, "visual.json"), 'w', encoding='utf-8') as f:
            json.dump(vis_json, f, indent=2)
        pages['pageOrder'].append(rpid)
        overrides.extend([
            f'<Override PartName="/Report/definition/pages/{rpid}/page.json" ContentType="application/json"/>',
            f'<Override PartName="/Report/definition/pages/{rpid}/visuals/{vis_json["name"]}/visual.json" ContentType="application/json"/>'
        ])

    save_pages_json(tmp, pages)
    for ov in overrides:
        if ov not in ct:
            ct = ct.replace("</Types>", ov + "\n</Types>")
    with open(ct_path, 'w', encoding='utf-8') as f:
        f.write(ct)
    remove_security_bindings(tmp)
    repack_pbix(tmp, args.output or args.pbix)
    print(f"Added ranking pages. Output: {args.output or args.pbix}")

# ─────────────────────────────────────────────
# COMMAND: rebuild-pbix
# ─────────────────────────────────────────────
def cmd_rebuild_pbix(args):
    """Rebuild PBIX with standard 4-page layout (Sales, Customer, Region, Product)."""
    tmp = "_pbix_build"
    extract_pbix(args.pbix, tmp)

    sales_pid = None
    pages = load_pages_json(tmp)
    for pid in pages['pageOrder']:
        pg = json.load(open(os.path.join(tmp, f'Report/definition/pages/{pid}/page.json'), encoding='utf-8'))
        if 'Sales' in pg.get('displayName', ''):
            sales_pid = pid
            pg['width'] = 1280
            pg['height'] = 720
            pg['displayOption'] = 'FitToPage'
            with open(os.path.join(tmp, f'Report/definition/pages/{pid}/page.json'), 'w', encoding='utf-8') as f:
                json.dump(pg, f, indent=2)
            break

    if sales_pid:
        card_ps = {k: v for k, v in {
            "v001": (30, 5, 195, 75, 0), "v002": (235, 5, 195, 75, 1),
            "v003": (440, 5, 195, 75, 2), "v004": (645, 5, 195, 75, 3),
            "v005": (850, 5, 195, 75, 4), "v006": (1055, 5, 195, 75, 5),
        }.items() if os.path.exists(os.path.join(tmp, f'Report/definition/pages/{sales_pid}/visuals/{k}/visual.json'))}
        for vid, (x, y, w, h, z) in card_ps.items():
            vf = os.path.join(tmp, f'Report/definition/pages/{sales_pid}/visuals/{vid}/visual.json')
            if os.path.exists(vf):
                v = json.load(open(vf, encoding='utf-8'))
                v['position'] = {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z}
                with open(vf, 'w', encoding='utf-8') as f:
                    json.dump(v, f, indent=2)

    M, S = "Measure", "SalesTable"
    overrides = []

    def add_page(pid, display_name, visuals):
        nonlocal overrides
        pdir = os.path.join(tmp, f"Report/definition/pages/{pid}")
        os.makedirs(os.path.join(pdir, "visuals"), exist_ok=True)
        with open(os.path.join(pdir, "page.json"), 'w', encoding='utf-8') as f:
            json.dump({
                "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/1.0.0/schema.json",
                "name": pid, "displayName": display_name,
                "displayOption": "FitToPage", "height": 720, "width": 1280
            }, f, indent=2)
        for v in visuals:
            vdir = os.path.join(pdir, "visuals", v["name"])
            os.makedirs(vdir, exist_ok=True)
            v["$schema"] = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.8.0/schema.json"
            with open(os.path.join(vdir, "visual.json"), 'w', encoding='utf-8') as f:
                json.dump(v, f, indent=2)
            overrides.append(f'<Override PartName="/Report/definition/pages/{pid}/visuals/{v["name"]}/visual.json" ContentType="application/json"/>')
        overrides.append(f'<Override PartName="/Report/definition/pages/{pid}/page.json" ContentType="application/json"/>')
        pages['pageOrder'].append(pid)

    dp_visuals = []
    for vid, tbl, meas, title, x, y, w, h, z in [
        ("dpc1", M, "Customers Covered (Display)", "Total Customers", 25, 5, 295, 75, 0),
        ("dpc2", M, "Active MRs (Display)", "Active MRs", 330, 5, 295, 75, 1),
        ("dpc3", M, "Total Revenue (CR)", "Total Revenue", 635, 5, 295, 75, 2),
        ("dpc4", M, "Total QTY (Display)", "Total QTY", 940, 5, 295, 75, 3),
    ]: dp_visuals.append(make_card(vid, tbl, meas, title, x, y, w, h, z))
    for vid, title, x, y, w, h, z in [
        ("dps1", "Region", 25, 95, 400, 40, 4),
        ("dps2", "Business Unit", 435, 95, 400, 40, 5),
        ("dps3", "Segment", 845, 95, 400, 40, 6),
    ]: dp_visuals.append(make_slicer(vid, title, x, y, w, h, z, S, title))
    for vid, vtype, title, x, y, w, h, z, ce, cp, vt, vp, oc in [
        ("dph1", "columnChart", "Revenue by Customer (₹ Lakh)", 25, 150, 400, 260, 7, S, "Customer Name", M, "Revenue (Lakh Numeric)", 1),
        ("dph2", "columnChart", "QTY by Customer", 435, 150, 400, 260, 8, S, "Customer Name", M, "QTY (Lakh Numeric)", 1),
        ("dph3", "barChart", "Avg Final Rate by Customer (₹)", 845, 150, 400, 260, 9, S, "Customer Name", M, "Avg Final Rate Numeric", 2),
    ]: dp_visuals.append(make_chart(vid, vtype, title, x, y, w, h, z, ce, cp, vt, vp, oc))
    for vid, vtype, title, x, y, w, h, z, ce, cp, vt, vp, oc in [
        ("dph4", "barChart", "Revenue by Clinic (₹ Lakh)", 25, 430, 600, 275, 10, S, "Clinic", M, "Revenue (Lakh Numeric)", 2),
        ("dph5", "barChart", "Avg PTR Rate by Customer (₹)", 635, 430, 600, 275, 11, S, "Customer Name", M, "Avg PTR Rate Numeric", 2),
    ]: dp_visuals.append(make_chart(vid, vtype, title, x, y, w, h, z, ce, cp, vt, vp, oc))

    to_visuals = []
    for vid, tbl, meas, title, x, y, w, h, z in [
        ("toc1", M, "Total Revenue (CR)", "Total Revenue", 25, 5, 295, 75, 0),
        ("toc2", M, "Total QTY (Display)", "Total QTY", 330, 5, 295, 75, 1),
        ("toc3", M, "Customers Covered (Display)", "Customers Covered", 635, 5, 295, 75, 2),
        ("toc4", M, "Active MRs (Display)", "Active MRs", 940, 5, 295, 75, 3),
    ]: to_visuals.append(make_card(vid, tbl, meas, title, x, y, w, h, z))
    for vid, title, x, y, w, h, z in [
        ("tos1", "Region", 25, 95, 400, 40, 4),
        ("tos2", "Business Unit", 435, 95, 400, 40, 5),
        ("tos3", "Sales Rep Name", 845, 95, 400, 40, 6),
    ]: to_visuals.append(make_slicer(vid, title, x, y, w, h, z, S, title))
    for vid, vtype, title, x, y, w, h, z, ce, cp, vt, vp, oc in [
        ("toh1", "columnChart", "Revenue by Region (₹ Cr)", 25, 150, 400, 260, 7, S, "Region", M, "Revenue (Cr Numeric)", 1),
        ("toh2", "columnChart", "QTY by Region", 435, 150, 400, 260, 8, S, "Region", M, "Total QTY", 1),
        ("toh3", "barChart", "Customers by Region", 845, 150, 400, 260, 9, S, "Region", M, "Customers Covered", 2),
    ]: to_visuals.append(make_chart(vid, vtype, title, x, y, w, h, z, ce, cp, vt, vp, oc))
    for vid, vtype, title, x, y, w, h, z, ce, cp, vt, vp, oc in [
        ("toh4", "barChart", "Revenue by Business Unit (₹ Cr)", 25, 430, 400, 275, 10, S, "Business Unit", M, "Revenue (Cr Numeric)", 2),
        ("toh5", "barChart", "QTY by Business Unit", 435, 430, 400, 275, 11, S, "Business Unit", M, "Total QTY", 2),
        ("toh6", "barChart", "Revenue by Sales Rep Name (₹ Lakh)", 845, 430, 400, 275, 12, S, "Sales Rep Name", M, "Revenue (Lakh Numeric)", 2),
    ]: to_visuals.append(make_chart(vid, vtype, title, x, y, w, h, z, ce, cp, vt, vp, oc))

    pa_visuals = []
    for vid, tbl, meas, title, x, y, w, h, z in [
        ("pac1", M, "Total Products (Display)", "Total Products", 25, 5, 295, 75, 0),
        ("pac2", M, "Total Revenue (CR)", "Total Revenue", 330, 5, 295, 75, 1),
        ("pac3", M, "Avg Final Rate (Display)", "Avg Final Rate", 635, 5, 295, 75, 2),
        ("pac4", M, "Avg PTR Rate (Display)", "Avg PTR Rate", 940, 5, 295, 75, 3),
    ]: pa_visuals.append(make_card(vid, tbl, meas, title, x, y, w, h, z))
    for vid, title, x, y, w, h, z in [
        ("pas1", "Business Unit", 25, 95, 400, 40, 4),
        ("pas2", "Region", 435, 95, 400, 40, 5),
        ("pas3", "Segment", 845, 95, 400, 40, 6),
    ]: pa_visuals.append(make_slicer(vid, title, x, y, w, h, z, S, title))
    for vid, vtype, title, x, y, w, h, z, ce, cp, vt, vp, oc in [
        ("pah1", "barChart", "Revenue by Product (₹ Lakh)", 25, 150, 400, 260, 7, S, "Product", M, "Revenue (Lakh Numeric)", 2),
        ("pah2", "columnChart", "QTY by Product (Lakh)", 435, 150, 400, 260, 8, S, "Product", M, "QTY (Lakh Numeric)", 1),
        ("pah3", "barChart", "Avg Final Rate by Product (₹)", 845, 150, 400, 260, 9, S, "Product", M, "Avg Final Rate Numeric", 2),
    ]: pa_visuals.append(make_chart(vid, vtype, title, x, y, w, h, z, ce, cp, vt, vp, oc))
    for vid, vtype, title, x, y, w, h, z, ce, cp, vt, vp, oc in [
        ("pah4", "barChart", "Avg PTR Rate by Product (₹)", 25, 430, 400, 275, 10, S, "Product", M, "Avg PTR Rate Numeric", 2),
        ("pah5", "columnChart", "Revenue by Business Unit (₹ Cr)", 435, 430, 400, 275, 11, S, "Business Unit", M, "Revenue (Cr Numeric)", 1),
        ("pah6", "lineChart", "Revenue Trend by Month (₹ Lakh)", 845, 430, 400, 275, 12, S, "Month", M, "Revenue (Lakh Numeric)", 0),
    ]: pa_visuals.append(make_chart(vid, vtype, title, x, y, w, h, z, ce, cp, vt, vp, oc))

    add_page(uuid.uuid4().hex[:20], "Customer Performance", dp_visuals)
    add_page(uuid.uuid4().hex[:20], "Region Overview", to_visuals)
    add_page(uuid.uuid4().hex[:20], "Product Analysis", pa_visuals)

    save_pages_json(tmp, pages)
    ct_path = os.path.join(tmp, "[Content_Types].xml")
    with open(ct_path, 'r', encoding='utf-8') as f:
        ct = f.read()
    for ov in overrides:
        if ov not in ct:
            ct = ct.replace("</Types>", ov + "\n</Types>")
    remove_security_bindings(tmp)
    for name in ["SecurityBindings", "SecurityBindings.xml"]:
        p = os.path.join(tmp, name)
        if os.path.isfile(p): os.remove(p)
        elif os.path.isdir(p): shutil.rmtree(p)
    with open(ct_path, 'w', encoding='utf-8') as f:
        f.write(ct)
    repack_pbix(tmp, args.output or args.pbix)
    print(f"Rebuilt PBIX with standard 4-page layout. Output: {args.output or args.pbix}")

# ─────────────────────────────────────────────
# COMMAND: overhaul (unified: v6 -> vX)
# ─────────────────────────────────────────────
def cmd_overhaul(args):
    """Apply all fixes (subtitle off, KPI styling, ranking pages, safe repack)."""
    tmp = f"_overhaul_v{args.version}_tmp"
    extract_pbix(args.pbix, tmp)
    pages = load_pages_json(tmp)
    ct_path = os.path.join(tmp, "[Content_Types].xml")
    with open(ct_path, 'r', encoding='utf-8') as f:
        ct = f.read()

    version = args.version

    for pid in list(pages['pageOrder']):
        for vid, v_json, vf in iter_visuals(tmp, pid):
            du = DISPLAY_UNITS_NONE if version <= 10 else DISPLAY_UNITS_THOUSANDS
            if version >= 13:
                du = DISPLAY_UNITS_THOUSANDS
            fix_visual(v_json, disable_subtitle=True, fix_kpi=True, display_units=du)
            if version >= 12:
                qs = v_json['visual'].get('query', {}).get('queryRegion', {})
                for shelf in qs.values():
                    for p in shelf.get('projections', []):
                        f = p.get('field', {})
                        if 'Measure' in f:
                            prop = f['Measure'].get('Property')
                            if prop == 'Total QTY':
                                f['Measure']['Property'] = 'QTY (Lakh Numeric)'
                                p['queryRef'] = 'QTY (Lakh Numeric)'
            write_visual(vf, v_json)

    # Remove old ranking pages (v11+)
    if version >= 11:
        keep_pids = []
        for pid in pages['pageOrder']:
            pname = get_page_name(tmp, pid)
            if pname not in ("Customer Performance Ranking", "Sales Performance & Splits", "Customer Ranking"):
                keep_pids.append(pid)
        pages['pageOrder'] = keep_pids

    # Add ranking pages (v7+)
    if version >= 7:
        ranking_configs = [
            ("Customer", "Customer Performance Ranking", uuid.uuid4().hex[:20],
             [("Customer Name (Combine Multiple)", "Customer Name")],
             [("Measure", "Total QTY Indian"), ("Measure", "Total Revenue Indian")],
             ("Measure", "Total Revenue")),
            ("MR", "Sales Performance & Splits", uuid.uuid4().hex[:20],
             [("SalesTable", "Sales Rep Name")],
             [("Measure", "Total Revenue (CR)"), ("Measure", "Customer Sales %"),
              ("Measure", "Partner Sales %"), ("Measure", "Sales Performance Detail")],
             ("Measure", "Total Revenue")),
        ]
        overrides = []
        for rname, rdisp, rpid, rcols, rmeas, rsort in ranking_configs:
            page_json, vis_json = create_ranking_page(rname, rdisp, rpid, rcols, rmeas, rsort)
            pdir = os.path.join(tmp, f"Report/definition/pages/{rpid}")
            vdir = os.path.join(pdir, "visuals", vis_json['name'])
            os.makedirs(vdir, exist_ok=True)
            with open(os.path.join(pdir, "page.json"), 'w', encoding='utf-8') as f:
                json.dump(page_json, f, indent=2)
            with open(os.path.join(vdir, "visual.json"), 'w', encoding='utf-8') as f:
                json.dump(vis_json, f, indent=2)
            pages['pageOrder'].append(rpid)
            overrides.extend([
                f'<Override PartName="/Report/definition/pages/{rpid}/page.json" ContentType="application/json"/>',
                f'<Override PartName="/Report/definition/pages/{rpid}/visuals/{vis_json["name"]}/visual.json" ContentType="application/json"/>'
            ])
        save_pages_json(tmp, pages)
        for ov in overrides:
            if ov not in ct:
                ct = ct.replace("</Types>", ov + "\n</Types>")

    remove_security_bindings(tmp)
    with open(ct_path, 'w', encoding='utf-8') as f:
        f.write(ct)
    repack_pbix(tmp, args.output or args.pbix)
    print(f"Overhaul v{version} completed. Output: {args.output or args.pbix}")

# ─────────────────────────────────────────────
# COMMAND: swap-datamodel
# ─────────────────────────────────────────────
def cmd_swap_dm(args):
    """Replace DataModel in target PBIX with DataModel from source PBIX."""
    tmp = "_swap_dm_tmp"
    output = args.output or args.target

    # Extract target
    extract_pbix(args.target, tmp)

    # Get DataModel from source
    src_dm = None
    src_dm_path = None
    with zipfile.ZipFile(args.source, 'r') as z:
        for n in z.namelist():
            if n.endswith('DataModel') and 'SecurityBindings' not in n:
                src_dm = z.read(n)
                src_dm_path = n
                break
    if not src_dm:
        print(f"DataModel not found in {args.source}")
        return

    # Replace DataModel in target
    replaced = False
    for root, dirs, files in os.walk(tmp):
        for fn in files:
            if fn == "DataModel":
                fp = os.path.join(root, fn)
                with open(fp, 'wb') as f:
                    f.write(src_dm)
                replaced = True
                print(f"Replaced DataModel: {fp} ({len(src_dm)} bytes)")

    if not replaced:
        print("DataModel not found in target")
        return

    remove_security_bindings(tmp)
    repack_pbix(tmp, output)
    print(f"DataModel swapped. Output: {output}")

# ─────────────────────────────────────────────
# COMMAND: save (Ctrl+S)
# ─────────────────────────────────────────────
def cmd_save(args):
    """Send Ctrl+S to Power BI Desktop to save."""
    success = send_ctrl_s(args.window_title)
    if success:
        print("Ctrl+S sent successfully.")
    else:
        print("Could not find Power BI window.")

# ─────────────────────────────────────────────
# COMMAND: backup
# ─────────────────────────────────────────────
def cmd_backup(args):
    """Create a .BACKUP copy of the PBIX file."""
    backup_path = args.pbix + ".BACKUP"
    shutil.copy2(args.pbix, backup_path)
    print(f"Backup created: {backup_path}")

# ─────────────────────────────────────────────
# V16 ENHANCED VISUAL FIXES
# ─────────────────────────────────────────────
def fix_auto_subtitle(v_json):
    v = v_json.get('visual', {})
    objs = v.setdefault('objects', {})
    # 1) Set objects.subTitle show=false (for theme-controlled auto-subtitle)
    objs.setdefault('subTitle', [{}])[0].setdefault('properties', {})['show'] = L("false")
    # 2) Set subtitle text to empty string (brute force — hide the text itself)
    objs['subTitle'][0]['properties']['text'] = L("''")
    objs['subTitle'][0]['properties']['fontSize'] = L("1D")
    # 3) Set autoSubtitle=false in general section (Format pane)
    objs.setdefault('general', [{}])[0].setdefault('properties', {})['autoSubtitle'] = L("false")
    return v_json

def add_desc_sort(v_json):
    v = v_json.get('visual', {})
    vtype = v.get('visualType', '')
    qs = v.get('query', {}).get('queryRegion', {})

    chart_types = ('barChart', 'columnChart', 'lineChart', 'stackedBarChart',
                   'stackedColumnChart', 'ribbonChart', 'waterfallChart', 'areaChart')
    if vtype not in chart_types:
        return False

    category = qs.get('Category', {})
    if not category.get('projections'):
        return False

    for proj in category['projections']:
        f = proj.get('field', {})
        for ftype, finfo in f.items():
            prop = finfo.get('Property', '')
            if 'Month' in prop or 'Date' in prop:
                return False

    value_field = None
    for shelf_name in ('Y', 'ColumnValues', 'Values'):
        shelf = qs.get(shelf_name, {})
        projs = shelf.get('projections', [])
        if projs and 'field' in projs[0]:
            value_field = projs[0]['field']
            break

    if not value_field:
        return False

    category['orderBy'] = [{"direction": 2, "field": value_field}]
    return True

def fix_visual_v16(v_json):
    v_json = fix_visual(v_json, disable_subtitle=True, fix_kpi=True, display_units=DISPLAY_UNITS_THOUSANDS)
    v_json = fix_auto_subtitle(v_json)
    add_desc_sort(v_json)
    return v_json

# ─────────────────────────────────────────────
# V16 DATAMODEL INJECTION
# ─────────────────────────────────────────────
def read_dm_schema(tmp):
    dm_path = os.path.join(tmp, 'DataModelSchema')
    if os.path.exists(dm_path):
        with open(dm_path, 'r', encoding='utf-8-sig') as f:
            return dm_path, json.load(f)
    for root, dirs, files in os.walk(tmp):
        for fn in files:
            if fn == "DataModelSchema":
                dm_path = os.path.join(root, fn)
                with open(dm_path, 'r', encoding='utf-8-sig') as f:
                    return dm_path, json.load(f)
    return None, None

def write_dm_schema(dm_path, schema):
    with open(dm_path, 'w', encoding='utf-8') as f:
        json.dump(schema, f, indent=2)

def find_table(schema, name):
    for t in schema.get('model', {}).get('tables', []):
        if t.get('name') == name:
            return t
    return None

def find_measure(schema, name):
    for t in schema.get('model', {}).get('tables', []):
        for m in t.get('measures', []):
            if m.get('name') == name:
                return m
    return None

def inject_canonical_customer_column(schema):
    sales_table = find_table(schema, 'SalesTable')
    if not sales_table:
        return False

    for c in sales_table.get('columns', []):
        if c.get('name') == 'Canonical Customer Name':
            return False

    col = {
        "type": 2,
        "name": "Canonical Customer Name",
        "dataType": "string",
        "isNullable": True,
        "isHidden": False,
        "expression": (
            "VAR CombineName = LOOKUPVALUE(\n"
            "    'Customer Name (Combine Multiple)'[Customer Name],\n"
            "    'Customer Name (Combine Multiple)'[Customer ID],\n"
            "    SalesTable[Customer ID]\n"
            ")\n"
            "RETURN IF(ISBLANK(CombineName), SalesTable[Customer Name], CombineName)"
        ),
        "annotations": [{"name": "PBI_ResultType", "value": "{\"ResultType\":\"Scalar\"}"}]
    }
    sales_table['columns'].append(col)
    return True

def inject_v16_measures(schema):
    model = schema.get('model', {})
    measure_table = find_table(schema, 'Measure')
    if not measure_table:
        return

    existing_measures = {m['name'] for m in measure_table.get('measures', [])}

    new_measures = {
        "Dynamic Growth % (Corrected)": (
            "VAR CurrentValue = [Revenue (Cr Numeric)]\n"
            "VAR PrevValue = CALCULATE(\n"
            "    [Revenue (Cr Numeric)],\n"
            "    DATEADD('Date'[Date], -1, MONTH)\n"
            ")\n"
            "RETURN DIVIDE(CurrentValue - PrevValue, PrevValue)"
        ),
        "QoQ Growth %": (
            "VAR CurrentQ = [Revenue (Cr Numeric)]\n"
            "VAR PrevQ = CALCULATE(\n"
            "    [Revenue (Cr Numeric)],\n"
            "    DATEADD('Date'[Date], -3, MONTH)\n"
            ")\n"
            "RETURN DIVIDE(CurrentQ - PrevQ, PrevQ)"
        ),
        "YoY Growth %": (
            "VAR CurrentVal = [Revenue (Cr Numeric)]\n"
            "VAR PrevVal = CALCULATE(\n"
            "    [Revenue (Cr Numeric)],\n"
            "    SAMEPERIODLASTYEAR('Date'[Date])\n"
            ")\n"
            "RETURN DIVIDE(CurrentVal - PrevVal, PrevVal)"
        ),
        "FYTD Revenue": (
            "VAR FYStart = STARTOFYEAR('Date'[Date], \"03-31\")\n"
            "RETURN CALCULATE(\n"
            "    [Total Revenue],\n"
            "    DATESYTD('Date'[Date], \"03-31\")\n"
            ")"
        ),
        "Trailing 12M Revenue": (
            "CALCULATE(\n"
            "    [Total Revenue],\n"
            "    DATESINPERIOD('Date'[Date], MAX('Date'[Date]), -12, MONTH)\n"
            ")"
        ),
        "Customer Channel Revenue": (
            "CALCULATE(\n"
            "    [Total Revenue],\n"
            "    'Customer Name (Combine Multiple)'\n"
            ")"
        ),
        "Partner Channel Revenue": (
            "CALCULATE(\n"
            "    [Total Revenue],\n"
            "    Partner_Hospital\n"
            ")"
        ),
        "Other Channel Revenue": (
            "[Total Revenue] - [Customer Channel Revenue] - [Partner Channel Revenue]"
        ),
        "Channel Type": (
            "SWITCH(TRUE(),\n"
            "    NOT ISEMPTY('Customer Name (Combine Multiple)'), \"Customer\",\n"
            "    NOT ISEMPTY(Partner_Hospital), \"Partner/Hospital\",\n"
            "    \"Other\"\n"
            ")"
        ),
        "Sales Rep Efficiency": (
            "DIVIDE([Total QTY], [Active MRs])"
        ),
        "Revenue Per Customer": (
            "DIVIDE([Total Revenue], [Customers Covered])"
        ),
    }

    ms = measure_table.setdefault('measures', [])
    added = 0
    for name, expr in new_measures.items():
        if name not in existing_measures:
            ms.append({
                "name": name, "expression": expr, "formatString": "#,##0.00",
                "isHidden": False
            })
            added += 1
    return added

# ─────────────────────────────────────────────
# V16 NEW PAGE BUILDERS
# ─────────────────────────────────────────────
OFF_CANVAS_X = 1290
SLICER_W = 170
SLICER_H = 50

def add_visual_to_page(page_dir, vid, v_json, ct, overrides):
    vdir = os.path.join(page_dir, 'visuals', vid)
    os.makedirs(vdir, exist_ok=True)
    v_json['name'] = vid
    v_json['$schema'] = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.8.0/schema.json"
    with open(os.path.join(vdir, 'visual.json'), 'w', encoding='utf-8') as f:
        json.dump(v_json, f, indent=2)
    ov = f'<Override PartName="/Report/definition/pages/{os.path.basename(page_dir)}/visuals/{vid}/visual.json" ContentType="application/json"/>'
    if ov not in ct:
        ct = ct.replace("</Types>", ov + "\n</Types>")
    overrides.append(ov)
    return ct

def get_page_pid_by_name(tmp, name):
    pages = load_pages_json(tmp)
    for pid in pages['pageOrder']:
        pn = get_page_name(tmp, pid)
        if pn == name:
            return pid
    return None

def add_offcanvas_slicers(tmp, ct, page_pid, suffix):
    M, S = "Measure", "SalesTable"
    slicer_defs = [
        ("sl_ter" + suffix, "Region", 2, S, 'Region'),
        ("sl_div" + suffix, "Business Unit", 56, S, 'Business Unit'),
        ("sl_rt" + suffix, "Segment", 110, S, 'Segment'),
        ("sl_mr" + suffix, "Sales Rep Name", 164, S, 'Sales Rep Name'),
    ]
    page_dir = os.path.join(tmp, f'Report/definition/pages/{page_pid}')
    for z, (vid, title, y, ent, prop) in enumerate(slicer_defs):
        sv = make_slicer(vid, title, OFF_CANVAS_X, y, SLICER_W, SLICER_H, z, ent, prop)
        sv['visual']['visualContainerObjects']['title'][0]['properties']['show'] = L("true")
        ct = add_visual_to_page(page_dir, vid, sv, ct, [])
    return ct

def make_explanation_box(vid, text, x, y, w, h, z):
    return {
        "name": vid,
        "position": {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z},
        "visual": {
            "visualType": "textbox",
            "objects": {
                "paragraphs": [{
                    "textRuns": [{
                        "textStyle": "normal",
                        "value": text
                    }]
                }]
            },
            "visualContainerObjects": {
                "background": [{"properties": {"color": {"solid": {"color": L("'#F5F5F5'")}}, "transparency": L("0D")}}],
                "border": [{"properties": {"show": L("true"), "radius": L("6D")}}]
            },
            "drillFilterOtherVisuals": False
        }
    }

def make_kpi_card(vid, val_text, title, x, y, w, h, z, color="#17614A"):
    return {
        "name": vid,
        "position": {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z},
        "visual": {
            "visualType": "card",
            "query": {"queryRegion": {"Values": {"projections": [
                {"field": meas_field("Measure", val_text), "queryRef": val_text, "nativeQueryRef": val_text}
            ]}}},
            "objects": {
                "categoryLabels": [{"properties": {"show": L("false")}}],
                "labels": [{"properties": {
                    "fontSize": L(KPI_VALUE_SIZE),
                    "color": {"solid": {"color": L(f"'{color}'")}},
                    "labelDisplayUnits": L("0D"), "labelPrecision": L("0L")
                }}],
                "wordWrap": [{"properties": {"show": L("true")}}]
            },
            "visualContainerObjects": {
                "title": [{"properties": {
                    "text": L(f"'{title}'"), "show": L("true"),
                    "titleWrap": L("true"),
                    "fontColor": {"solid": {"color": L("'#333333'")}},
                    "background": {"solid": {"color": L(f"'{TITLE_BG}'")}},
                    "alignment": L("'center'"), "fontSize": L(KPI_TITLE_SIZE)
                }}],
                "visualHeader": [{"properties": {"show": L("true"), "transparency": L("100D")}}],
                "border": [{"properties": {"show": L("true"), "radius": L(CONTAINER_RADIUS), "width": L("1D"), "color": {"solid": {"color": L(f"'{BORDER_COLOR}'")}}}}],
                "dropShadow": [{"properties": {"show": L("true")}}]
            },
            "drillFilterOtherVisuals": True
        }
    }

# ─────────────────────────────────────────────
# COMMAND: build-v16
# ─────────────────────────────────────────────
def cmd_build_v16(args):
    """Build v16 from v15: ONLY modify existing visuals + patch theme — NO new visuals/pages from scratch.
    New visuals created from scratch in JSON cause 'Missing_References' errors because Power BI Desktop
    can't resolve field references in manually-crafted visual JSON against the binary DataModel."""
    SRC = args.pbix
    OUT = args.output or SRC.replace('v15', 'v16').replace('.pbix', '.v16.pbix')
    TMP = "_build_v16_tmp"
    M = "Measure"

    print("=== Step 1: Extract v15 ===")
    extract_pbix(SRC, TMP)
    pages = load_pages_json(TMP)
    ct_path = os.path.join(TMP, "[Content_Types].xml")
    with open(ct_path, 'r', encoding='utf-8') as f:
        ct = f.read()
    print("  Extracted.", flush=True)

    print("=== Step 2: Remove empty 'Page 1' ===")
    keep_pids = []
    for pid in pages['pageOrder']:
        pname = get_page_name(TMP, pid)
        if pname == "Page 1":
            pdir = os.path.join(TMP, f'Report/definition/pages/{pid}')
            if os.path.exists(pdir):
                shutil.rmtree(pdir)
        else:
            keep_pids.append(pid)
    pages['pageOrder'] = keep_pids
    save_pages_json(TMP, pages)
    print("  Removed empty page(s).", flush=True)

    print("=== Step 3: Fix auto-subtitles + Descending sort + KPI sizing ===")
    visual_count = 0
    for pid in list(pages['pageOrder']):
        for vid, v_json, vf in iter_visuals(TMP, pid):
            fix_visual_v16(v_json)
            write_visual(vf, v_json)
            visual_count += 1
    print(f"  Fixed {visual_count} existing visuals.", flush=True)

    print("=== Step 4: Fix ranking page (blank names + sort) ===")
    for pid in list(pages['pageOrder']):
        pname = get_page_name(TMP, pid)
        if pname == "Customer Performance Ranking":
            for vid, v_json, vf in iter_visuals(TMP, pid):
                qs = v_json['visual'].get('query', {}).get('queryRegion', {})
                projs = qs.get('Values', {}).get('projections', [])
                # Switch from Combine table -> SalesTable.Customer Name to fix blank names
                for p in projs:
                    col_f = p.get('field', {}).get('Column', {})
                    sr = col_f.get('Expression', {}).get('SourceRef', {}).get('Entity', '')
                    if 'Combine' in sr and col_f.get('Property') == 'Customer Name':
                        col_f['Expression']['SourceRef']['Entity'] = 'SalesTable'
                        p['queryRef'] = 'SalesTable.Customer Name'
                # Add hidden Revenue (Cr Numeric) sort column
                if not any(p.get('queryRef') == 'Revenue (Cr Numeric)' for p in projs):
                    projs.insert(0, {
                        "field": meas_field(M, "Revenue (Cr Numeric)"),
                        "queryRef": "Revenue (Cr Numeric)", "isHidden": True
                    })
                qs['Values']['orderBy'] = [
                    {"direction": 2, "field": meas_field(M, "Revenue (Cr Numeric)")}
                ]
                write_visual(vf, v_json)
            print("  Customer Performance Ranking: switched to SalesTable.Customer Name + Revenue sort.", flush=True)

    print("=== Step 5: Patch Fluent2 theme (disable subTitle at theme level) ===")
    theme_patched = False
    for root, dirs, files in os.walk(TMP):
        for fn in files:
            if 'Fluent' in fn and fn.endswith('.json'):
                fp = os.path.join(root, fn)
                theme = json.load(open(fp, encoding='utf-8-sig'))
                vs = theme.get('visualStyles', {})
                wc = vs.get('*', {})
                chart_types = ['*', 'barChart', 'columnChart', 'lineChart', 'areaChart',
                               'stackedBarChart', 'stackedColumnChart', 'card', 'slicer',
                               'tableEx', 'ribbonChart', 'waterfallChart', 'pieChart',
                               'stackedAreaChart', 'clusteredColumnChart', 'comboChart']
                for vtype in chart_types:
                    vval = wc.get(vtype)
                    if vval is None:
                        continue
                    sub = vval.get('subTitle')
                    if sub and len(sub) > 0:
                        if sub[0].get('show', True) == True:
                            sub[0]['show'] = False
                            theme_patched = True
                if theme_patched:
                    with open(fp, 'w', encoding='utf-8') as f:
                        json.dump(theme, f, indent=2)
                    print(f"  Fluent2 theme patched: subTitle.show=False for {len([t for t in chart_types if wc.get(t)])} visual types.", flush=True)
                break
        if theme_patched:
            break
    if not theme_patched:
        print("  Theme file: no subtitle settings to patch (or file not found).", flush=True)

    print("=== Step 6: Finalize - Security + Repack ===")
    remove_security_bindings(TMP)
    repack_pbix(TMP, OUT)
    print(f"  Output: {OUT}", flush=True)

    print("=== Verification ===")
    cmd_verify(argparse.Namespace(pbix=OUT))

    print("")
    print("=" * 60)
    print("  v16 BUILD COMPLETE — existing visuals only")
    print("=" * 60)
    print("")
    print("  Changes made:")
    print("  \u2022 Subtitle disabled (theme-level + visual-level)")
    print("  \u2022 Descending sort on all charts (except Month lines)")
    print("  \u2022 KPI: Title 10pt, Value 14pt (fit 75px boxes)")
    print("  \u2022 Ranking sort by Revenue (Cr Numeric)")
    print("  \u2022 Empty 'Page 1' removed")
    print("")
    print("  NEXT: Open v16.pbix in Power BI Desktop.")
    print("  It will look IDENTICAL to v15 but without subtitles and with correct sort/KPI sizing.")
    print("")
    print("  To add NEW features (Exec Summary, filter sidebar, menu button), you MUST create")
    print("  them manually in Power BI Desktop and then I can help style them.")
    print("")

# ─────────────────────────────────────────────
# V17 — COMMERCIAL UPGRADE
# ─────────────────────────────────────────────
def make_chart_v2(vid, vtype, title, x, y, w, h, z,
                  cat_ent, cat_prop, val_ent, val_prop,
                  orient=1, sort_desc=False):
    """Chart with Column ref for category (not Aggregation) — mirrors PBI Desktop format."""
    qs = {"Category": {"projections": [{
        "field": col_field(cat_ent, cat_prop),
        "queryRef": f"{cat_ent}.{cat_prop}",
        "nativeQueryRef": cat_prop, "active": True
    }]}, "Y": {"projections": [{
        "field": meas_field(val_ent, val_prop),
        "queryRef": "Val0", "nativeQueryRef": val_prop
    }]}}
    if sort_desc:
        qs["Category"]["orderBy"] = [{"direction": 2, "field": meas_field(val_ent, val_prop)}]
    obj = {
        "name": vid,
        "position": {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z},
        "visual": {
            "visualType": vtype,
            "query": {"queryRegion": qs},
            "objects": {
                "general": [{"properties": {"autoSubtitle": L("false")}}],
                "legend": [{"properties": {"show": L("false")}}],
                "labels": [{"properties": {"show": L("false")}}],
                "categoryAxis": [{"properties": {"labelColor": {"solid": {"color": L("'#333333'")}}, "showAxisTitle": L("false"), "fontSize": L("10D"), "bold": L("false"), "concatenateLabels": L("true")}}],
                "valueAxis": [{"properties": {"labelColor": {"solid": {"color": L("'#333333'")}}, "showAxisTitle": L("false"), "show": L("true"), "gridlineShow": L("false"), "fontSize": L("10D")}}],
                "subTitle": [{"properties": {"show": L("false"), "text": L("''"), "fontSize": L("1D")}}]
            },
            "visualContainerObjects": {
                "title": [{"properties": {"text": L(f"'{title}'"), "show": L("true"), "titleWrap": L("true"), "fontColor": {"solid": {"color": L("'#333333'")}}, "background": {"solid": {"color": L(f"'{TITLE_BG}'")}}, "alignment": L("'center'"), "fontSize": L(CHART_TITLE_SIZE)}}],
                "subTitle": [{"properties": {"show": L("false"), "text": L("''")}}],
                "visualHeader": [{"properties": {"show": L("true"), "transparency": L("100D")}}],
                "border": [{"properties": {"show": L("true"), "radius": L(BORDER_RADIUS), "width": L("1D"), "color": {"solid": {"color": L(f"'{BORDER_COLOR}'")}}}}],
                "dropShadow": [{"properties": {"show": L("true")}}]
            },
            "drillFilterOtherVisuals": True
        }
    }
    if vtype in ("barChart", "columnChart"):
        obj["visual"]["objects"]["general"][0]["properties"]["orientation"] = L(f"{orient}D")
    return obj

def make_nav_text(vid, text, x, y, z):
    return {
        "name": vid,
        "position": {"x": x, "y": y, "z": z, "height": 30, "width": 140, "tabOrder": z},
        "visual": {
            "visualType": "textbox",
            "objects": {
                "paragraphs": [{"textRuns": [{"textStyle": "normal", "value": text}]}]
            },
            "visualContainerObjects": {
                "background": [{"properties": {"color": {"solid": {"color": L("'#17614A'")}}, "transparency": L("0D")}}],
                "visualHeader": [{"properties": {"show": L("true"), "transparency": L("100D")}}]
            },
            "drillFilterOtherVisuals": False
        }
    }

def create_new_page(tmp, pid, display_name, ct):
    """Create a new blank page folder and return its page_dir path."""
    page_dir = os.path.join(tmp, f'Report/definition/pages/{pid}')
    os.makedirs(page_dir, exist_ok=True)
    page_json = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/1.0.0/schema.json",
        "name": pid, "displayName": display_name, "displayOption": "FitToPage", "height": 720, "width": 1280
    }
    with open(os.path.join(page_dir, 'page.json'), 'w', encoding='utf-8') as f:
        json.dump(page_json, f, indent=2)
    ov = f'<Override PartName="/Report/definition/pages/{pid}/page.json" ContentType="application/json"/>'
    if ov not in ct:
        ct = ct.replace("</Types>", ov + "\n</Types>")
    return page_dir, ct

def cmd_build_v17(args):
    """Build v17 from v16c: Executive Summary page + off-canvas slicers + navigation bar."""
    SRC = args.pbix
    OUT = args.output or SRC.replace('v16c', 'v17').replace('v16', 'v17').replace('.pbix', '.v17.pbix')
    TMP = "_build_v17_tmp"
    M, S = "Measure", "SalesTable"

    print("=== Step 1: Extract v16c ===")
    extract_pbix(SRC, TMP)
    pages = load_pages_json(TMP)
    with open(os.path.join(TMP, '[Content_Types].xml'), 'r', encoding='utf-8') as f:
        ct = f.read()
    print(f"  Extracted {len(pages['pageOrder'])} pages.", flush=True)

    print("=== Step 1b: Disable auto-subtitle in Fluent2 theme (global) ===")
    theme_path = os.path.join(TMP, 'Report/StaticResources/SharedResources/BaseThemes/Fluent2-CY26SU05.json')
    if os.path.isfile(theme_path):
        with open(theme_path, 'r', encoding='utf-8-sig') as f:
            theme = json.load(f)
        theme.setdefault('visualStyles', {}).setdefault('*', {}).setdefault('*', {})['general'] = [{'autoSubtitle': False}]
        with open(theme_path, 'w', encoding='utf-8-sig') as f:
            json.dump(theme, f, ensure_ascii=False, indent=2)
        print("  Fluent2 theme: general.autoSubtitle=False set globally.", flush=True)
    else:
        print("  Fluent2 theme not found — skipping.", flush=True)

    print("=== Step 2: Create Executive Summary page ===")
    es_pid = str(uuid.uuid4()).replace('-', '')[:20]
    es_dir, ct = create_new_page(TMP, es_pid, "Executive Summary", ct)

    kpi_w, kpi_h = 205, 75
    kpi_xs = [10 + i * (kpi_w + 6) for i in range(6)]
    kpis = [
        ("es_kpi1", "Total Revenue (CR)", "Total Revenue", kpi_xs[0], 10, kpi_w, kpi_h, 0),
        ("es_kpi2", "Total QTY (Display)", "Total QTY", kpi_xs[1], 10, kpi_w, kpi_h, 1),
        ("es_kpi3", "Customers Covered (Display)", "Customers Covered", kpi_xs[2], 10, kpi_w, kpi_h, 2),
        ("es_kpi4", "Total Products (Display)", "Total Products", kpi_xs[3], 10, kpi_w, kpi_h, 3),
        ("es_kpi5", "Active MRs (Display)", "Active MRs", kpi_xs[4], 10, kpi_w, kpi_h, 4),
        ("es_kpi6", "Avg Final Rate (Display)", "Avg Final Rate", kpi_xs[5], 10, kpi_w, kpi_h, 5),
    ]
    for vid, meas, title, x, y, w, h, z in kpis:
        cv = make_card(vid, M, meas, title, x, y, w, h, z)
        ct = add_visual_to_page(es_dir, vid, cv, ct, [])

    charts = [
        ("es_ch1", "lineChart", "Monthly Revenue Trend", 6, 94, 633, 308, 6, S, "Month", M, "Revenue (Cr Numeric)", False),
        ("es_ch2", "barChart", "Top Products by Revenue", 645, 94, 629, 308, 7, S, "Product", M, "Revenue (Lakh Numeric)", True),
        ("es_ch3", "barChart", "Top Customers by Revenue", 6, 412, 633, 300, 8, S, "Customer Name", M, "Revenue (Lakh Numeric)", True),
        ("es_ch4", "barChart", "Business Unit Performance", 645, 412, 629, 300, 9, S, "Business Unit", M, "Revenue (Lakh Numeric)", True),
    ]
    for vid, vtype, title, x, y, w, h, z, ce, cp, ve, vp, sd in charts:
        cv = make_chart_v2(vid, vtype, title, x, y, w, h, z, ce, cp, ve, vp, 1, sd)
        ct = add_visual_to_page(es_dir, vid, cv, ct, [])

    # Off-canvas slicers on Executive Summary
    ct = add_offcanvas_slicers(TMP, ct, es_pid, "_es")

    # Navigation bar (text labels at bottom of Executive Summary)
    nav_pages = [
        ("es_nav1", "Sales Dashboard"),
        ("es_nav2", "Customer Performance"),
        ("es_nav3", "Region Overview"),
        ("es_nav4", "Product Analysis"),
        ("es_nav5", "Customer Ranking"),
        ("es_nav6", "Sales Performance"),
    ]
    for i, (vid, label) in enumerate(nav_pages):
        nx = 6 + i * 210
        nv = make_nav_text(vid, label, nx, 690, 50 + i)
        ct = add_visual_to_page(es_dir, vid, nv, ct, [])

    pages['pageOrder'].insert(0, es_pid)
    save_pages_json(TMP, pages)
    print("  Executive Summary page created with KPIs, charts, slicers, and nav bar.", flush=True)

    print("=== Step 3: Add off-canvas slicers to existing pages ===")
    pids_by_name = {}
    for pid in pages['pageOrder']:
        pn = get_page_name(TMP, pid)
        pids_by_name[pn] = pid
    target_pages = ["Sales Dashboard", "Customer Performance", "Region Overview", "Product Analysis",
                     "Customer Performance Ranking", "Sales Performance & Splits"]
    suffixes = ["_sd", "_dp", "_to", "_pa", "_dr", "_mr"]
    for pname, suffix in zip(target_pages, suffixes):
        if pname in pids_by_name:
            ct = add_offcanvas_slicers(TMP, ct, pids_by_name[pname], suffix)
            print(f"  + slicers on {pname}", flush=True)

    print("=== Step 4: Force subtitle suppression on ALL visuals ===")
    for pid in pages['pageOrder']:
        pname = get_page_name(TMP, pid)
        for vid, v_json, vf in iter_visuals(TMP, pid):
            v_json = fix_auto_subtitle(v_json)
            v_json = fix_visual(v_json, disable_subtitle=True, fix_kpi=False, display_units=None, fix_aggregations=False)
            write_visual(vf, v_json)
    print("  Subtitles suppressed on ALL visuals.", flush=True)

    print("=== Step 5: Finalize - Security + Repack ===")
    remove_security_bindings(TMP)
    repack_pbix(TMP, OUT)
    print(f"  Output: {OUT}", flush=True)

    print("\n" + "=" * 60)
    print("  v17 BUILD COMPLETE — Commercial Upgrade")
    print("=" * 60)
    print("")
    print("  What's new:")
    print("  • Executive Summary page (page 1) — 6 KPIs, 4 charts")
    print("  • Off-canvas slicers on ALL pages (Region, Business Unit, Segment, Sales Rep Name)")
    print("  • Navigation bar at bottom of Executive Summary")
    print("")
    print("  NEXT: Open v17.pbix in Power BI Desktop to verify.")
    print("  If you see yellow triangles, open each chart and re-select the field in the Fields pane.")
    print("  To add: filter show/hide bookmarks + clickable nav buttons in Power BI Desktop.")
    print("")

# ─────────────────────────────────────────────
# COMMAND: measures inject-v16 (TOM via SSAS)
# ─────────────────────────────────────────────
def cmd_inject_v16(args):
    """Inject v16 model changes via SSAS (open v16.pbix in Power BI first)."""
    port = find_port_retry()
    print(f"Connected to SSAS port {port}")
    try:
        svr = connect_tom(port)
        db = svr.Databases[0]
        model = db.Model

        # 1. Add Canonical Customer Name calculated column
        sales_table = None
        for t in model.Tables:
            if t.Name == "SalesTable":
                sales_table = t
                break
        if sales_table:
            from Microsoft.AnalysisServices.Tabular import DataType
            existing = None
            for c in sales_table.Columns:
                if c.Name == "Canonical Customer Name":
                    existing = c
                    break
            if not existing:
                col = sales_table.Columns.AddNew()
                col.Name = "Canonical Customer Name"
                col.DataType = DataType.String
                col.IsNullable = True
                col.IsHidden = False
                col.Expression = (
                    "VAR CombineName = LOOKUPVALUE(\n"
                    "    'Customer Name (Combine Multiple)'[Customer Name],\n"
                    "    'Customer Name (Combine Multiple)'[Customer ID],\n"
                    "    SalesTable[Customer ID]\n"
                    ")\n"
                    "RETURN IF(ISBLANK(CombineName), SalesTable[Customer Name], CombineName)"
                )
                print("  Added Canonical Customer Name calculated column.")
            else:
                print("  Canonical Customer Name already exists.")

        # 2. New measures
        measure_table = None
        for t in model.Tables:
            if t.Name == "Measure":
                measure_table = t
                break
        if not measure_table:
            measure_table = model.Tables.AddNew()
            measure_table.Name = "Measure"
            dc = measure_table.Columns.AddNew()
            dc.Name = "Measure Placeholder"
            dc.DataType = DataType.Int64
            dc.IsNullable = True
            dc.SourceColumn = "Measure Placeholder"
            print("  Created Measure table")

        from Microsoft.AnalysisServices.Tabular import Measure as TOMeasure
        existing_names = set()
        for t in model.Tables:
            for m in t.Measures:
                existing_names.add(m.Name)

        new_measures = {
            "QoQ Growth %": ("VAR CurrentQ = [Revenue (Cr Numeric)] VAR PrevQ = CALCULATE([Revenue (Cr Numeric)], DATEADD('Date'[Date], -3, MONTH)) RETURN DIVIDE(CurrentQ - PrevQ, PrevQ)", "0.00%"),
            "YoY Growth %": ("VAR CurrentVal = [Revenue (Cr Numeric)] VAR PrevVal = CALCULATE([Revenue (Cr Numeric)], SAMEPERIODLASTYEAR('Date'[Date])) RETURN DIVIDE(CurrentVal - PrevVal, PrevVal)", "0.00%"),
            "FYTD Revenue": ("CALCULATE([Total Revenue], DATESYTD('Date'[Date], \"03-31\"))", "#,##0.00"),
            "Trailing 12M Revenue": ("CALCULATE([Total Revenue], DATESINPERIOD('Date'[Date], MAX('Date'[Date]), -12, MONTH))", "#,##0.00"),
            "Sales Rep Efficiency": ("DIVIDE([Total QTY], [Active MRs])", "#,##0"),
            "Revenue Per Customer": ("DIVIDE([Total Revenue], [Customers Covered])", "#,##0.00"),
            "Channel Type": ("SWITCH(TRUE(), NOT ISEMPTY('Customer Name (Combine Multiple)'), \"Customer\", NOT ISEMPTY(Partner_Hospital), \"Partner/Hospital\", \"Other\")", "0"),
        }

        for name, (expr, fmt) in new_measures.items():
            if name not in existing_names:
                nm = TOMeasure()
                nm.Name = name
                nm.Expression = expr
                nm.FormatString = fmt
                measure_table.Measures.Add(nm)
                print(f"  Added measure: {name}")

        # 3. Fix Dynamic Growth % to correct Date table reference
        for t in model.Tables:
            for m in t.Measures:
                if m.Name == "Dynamic Growth %" or m.Name == "Dynamic Growth % (Corrected)":
                    m.Expression = "VAR CurrentValue = [Revenue (Cr Numeric)] VAR PrevValue = CALCULATE([Revenue (Cr Numeric)], DATEADD('Date'[Date], -1, MONTH)) RETURN DIVIDE(CurrentValue - PrevValue, PrevValue)"
                    m.FormatString = "0.00%"
                    print(f"  Fixed measure: {m.Name}")

        model.SaveChanges()
        svr.Disconnect()
        print("\n  Model changes saved successfully.")
        print("  Close and reopen the PBIX in Power BI to see the new columns and measures.")
    except Exception as e:
        print(f"TOM error: {e}")

# ─────────────────────────────────────────────
# COMMAND: measures (TOM via SSAS)
# ─────────────────────────────────────────────
def cmd_list_measures(args):
    """List all measures from SSAS model (requires Power BI open)."""
    port = find_port_retry()
    print(f"Connected to port {port}")
    try:
        svr = connect_tom(port)
        db = svr.Databases[0]
        for table in db.Model.Tables:
            if table.Measures.Count > 0:
                print(f"\n[{table.Name}]")
                for m in table.Measures:
                    corrupted = " CORRUPTED" if "CORRUPT" in (m.Expression or "").upper() else ""
                    print(f"  {m.Name}{corrupted}")
        svr.Disconnect()
    except Exception as e:
        print(f"TOM error: {e}")

def cmd_add_measures(args):
    """Add/update measures in SSAS model."""
    port = find_port_retry()
    print(f"Connected to port {port}")
    try:
        svr = connect_tom(port)
        db = svr.Databases[0]
        model = db.Model

        # Ensure Measure table exists
        mt = None
        for t in model.Tables:
            if t.Name == "Measure":
                mt = t
                break
        if not mt:
            mt = model.Tables.AddNew()
            mt.Name = "Measure"
            # Add placeholder column
            from Microsoft.AnalysisServices.Tabular import DataType
            dc = mt.Columns.AddNew()
            dc.Name = "Measure Placeholder"
            dc.DataType = DataType.Int64
            dc.IsNullable = True
            dc.SourceColumn = "Measure Placeholder"
            print("Created Measure table")

        # Fix corruption first
        for t in model.Tables:
            for m in t.Measures:
                if "CORRUPT" in (m.Expression or "").upper():
                    m.Expression = "1"
                    print(f"Fixed corruption: {m.Name}")

        # Add/update measures
        measures_to_add = args.measures_list.split(',') if args.measures_list else list(BASE_MEASURES_DICT.keys())
        for mname in measures_to_add:
            if mname in BASE_MEASURES_DICT:
                mexpr = BASE_MEASURES_DICT[mname]
                existing = None
                for t in model.Tables:
                    for m in t.Measures:
                        if m.Name == mname:
                            existing = m
                            break
                if existing:
                    existing.Expression = mexpr
                    print(f"Updated: {mname}")
                else:
                    from Microsoft.AnalysisServices.Tabular import Measure as TOMeasure
                    nm = TOMeasure()
                    nm.Name = mname
                    nm.Expression = mexpr
                    mt.Measures.Add(nm)
                    print(f"Added: {mname}")

        model.SaveChanges()
        svr.Disconnect()
        print("Measures saved successfully.")
    except Exception as e:
        print(f"TOM error: {e}")

def cmd_fix_formats(args):
    """Fix measure format strings via SSAS."""
    port = find_port_retry()
    print(f"Connected to port {port}")
    try:
        svr = connect_tom(port)
        db = svr.Databases[0]
        model = db.Model

        PERCENTAGE_MEASURES = {"Customer Sales %", "Partner Sales %", "Dynamic Growth %", "Region Revenue %",
                               "Growth Category", "Growth Arrow", "Growth Color", "Sales Performance Detail"}

        for t in model.Tables:
            for m in t.Measures:
                if m.Name in PERCENTAGE_MEASURES:
                    m.FormatString = "0.00%"
                elif m.Name in ("Total Revenue Indian", "Total QTY Indian"):
                    m.FormatString = "0.00"
                elif any(d in m.Name for d in ("Display", "Indian")):
                    m.FormatString = "#,##0"
                elif any(r in m.Name for r in ("Rate", "Avg")):
                    m.FormatString = "#,##0.00"
                else:
                    m.FormatString = "#,##0"

        model.SaveChanges()
        svr.Disconnect()
        print("Format strings updated.")
    except Exception as e:
        print(f"TOM error: {e}")

def cmd_audit_dax(args):
    """Check for corrupted measures and list all."""
    port = find_port_retry()
    print(f"Connected to port {port}")
    try:
        svr = connect_tom(port)
        db = svr.Databases[0]
        model = db.Model
        corrupted = 0
        for t in model.Tables:
            for m in t.Measures:
                is_corrupt = "CORRUPT" in (m.Expression or "").upper()
                if is_corrupt:
                    print(f"CORRUPTED: [{t.Name}] {m.Name}")
                    m.Expression = "1"
                    corrupted += 1
                print(f"  [{t.Name}] {m.Name} = {m.Expression[:80]}...")
        if corrupted:
            model.SaveChanges()
            print(f"Fixed {corrupted} corrupted measures.")
        svr.Disconnect()
    except Exception as e:
        print(f"TOM error: {e}")

def cmd_export_reference(args):
    """Export all measures to a text file."""
    port = find_port_retry()
    out = args.output or "reference_measures.txt"
    try:
        svr = connect_tom(port)
        db = svr.Databases[0]
        model = db.Model
        with open(out, 'w', encoding='utf-8') as f:
            for t in model.Tables:
                if t.Measures.Count > 0:
                    f.write(f"\n{'='*60}\nTable: {t.Name}\n{'='*60}\n")
                    for m in t.Measures:
                        f.write(f"\nMeasure: {m.Name}\n")
                        f.write(f"Format: {m.FormatString}\n")
                        f.write(f"Expression:\n{m.Expression}\n")
        svr.Disconnect()
        print(f"Exported to {out}")
    except Exception as e:
        print(f"TOM error: {e}")

# ─────────────────────────────────────────────
# CLI ARGUMENT PARSER
# ─────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        prog='powerbi_automation',
        description='PowerBI Automation Suite — Unified CLI for PBIX manipulation',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""\
            Examples:
              python powerbi_automation.py analyze my.pbix
              python powerbi_automation.py fix-visuals my.pbix -o fixed.pbix --fix-aggs
              python powerbi_automation.py overhaul my.pbix -v 15 -o v15_output.pbix
              python powerbi_automation.py rebuild-pbix my.pbix -o rebuilt.pbix
              python powerbi_automation.py add-ranking-page my.pbix
              python powerbi_automation.py reorder-pages my.pbix "Sales Dashboard,Customer Performance"
              python powerbi_automation.py measures list
              python powerbi_automation.py save
        """)
    )
    parser.add_argument('--pbix', default=None, help='Default PBIX file path')

    sub = parser.add_subparsers(dest='command', required=True)

    # analyze
    p = sub.add_parser('analyze', help='Show full PBIX structure')
    p.add_argument('pbix'); p.set_defaults(func=cmd_analyze)

    # analyze-visuals
    p = sub.add_parser('analyze-visuals', help='Show visual types/titles/fields')
    p.add_argument('pbix'); p.set_defaults(func=cmd_analyze_visuals)

    # verify
    p = sub.add_parser('verify', help='Verify PBIX integrity')
    p.add_argument('pbix'); p.set_defaults(func=cmd_verify)

    # list-pages
    p = sub.add_parser('list-pages', help='List pages in PBIX')
    p.add_argument('pbix'); p.set_defaults(func=cmd_list_pages)

    # reorder-pages
    p = sub.add_parser('reorder-pages', help='Reorder pages by display names')
    p.add_argument('pbix'); p.add_argument('order', help='Comma-separated display names')
    p.add_argument('-o', '--output'); p.set_defaults(func=cmd_reorder_pages)

    # remove-pages
    p = sub.add_parser('remove-pages', help='Remove pages by display names')
    p.add_argument('pbix'); p.add_argument('names', help='Comma-separated display names')
    p.add_argument('-o', '--output'); p.set_defaults(func=cmd_remove_pages)

    # add-page
    p = sub.add_parser('add-page', help='Add a new blank page')
    p.add_argument('pbix'); p.add_argument('name', help='Page display name')
    p.add_argument('-o', '--output'); p.set_defaults(func=cmd_add_page)

    # fix-visuals
    p = sub.add_parser('fix-visuals', help='Fix all visuals (subtitles, KPI, display units)')
    p.add_argument('pbix'); p.add_argument('-o', '--output')
    p.add_argument('--display-units', default=None, help='Display units value (0D, 1L, etc)')
    p.add_argument('--fix-aggs', action='store_true', help='Fix aggregation-to-measure mappings')
    p.set_defaults(func=cmd_fix_visuals)

    # align-visuals
    p = sub.add_parser('align-visuals', help='Align visuals in grid layout')
    p.add_argument('pbix'); p.add_argument('-o', '--output')
    p.set_defaults(func=cmd_align_visuals)

    # add-ranking-page
    p = sub.add_parser('add-ranking-page', help='Add Customer/Sales Rep ranking pages')
    p.add_argument('pbix'); p.add_argument('-o', '--output')
    p.set_defaults(func=cmd_add_ranking_page)

    # rebuild-pbix
    p = sub.add_parser('rebuild-pbix', help='Rebuild with 4-page standard layout')
    p.add_argument('pbix'); p.add_argument('-o', '--output')
    p.set_defaults(func=cmd_rebuild_pbix)

    # overhaul
    p = sub.add_parser('overhaul', help='Apply vX overhaul (subtitle+KPI+ranking+repack)')
    p.add_argument('pbix'); p.add_argument('-v', '--version', type=int, default=15, help='Version (7-15)')
    p.add_argument('-o', '--output'); p.set_defaults(func=cmd_overhaul)

    # build-v16
    p = sub.add_parser('build-v16', help='Build v16 from v15: auto-subtitles, desc sort, canonical names, CEO dashboard')
    p.add_argument('pbix', help='v15 PBIX file')
    p.add_argument('--out', '-o', dest='output')
    p.set_defaults(func=cmd_build_v16)

    # build-v17
    p = sub.add_parser('build-v17', help='Build v17 from v16c: Executive Summary page, off-canvas slicers, navigation')
    p.add_argument('pbix', help='v16c PBIX file')
    p.add_argument('--out', '-o', dest='output')
    p.set_defaults(func=cmd_build_v17)

    # swap-datamodel
    p = sub.add_parser('swap-datamodel', help='Replace DataModel between PBIX files')
    p.add_argument('source', help='Source PBIX with correct DataModel')
    p.add_argument('target', help='Target PBIX to patch')
    p.add_argument('-o', '--output'); p.set_defaults(func=cmd_swap_dm)

    # save
    p = sub.add_parser('save', help='Send Ctrl+S to Power BI Desktop')
    p.add_argument('--window-title', help='Specific window title')
    p.set_defaults(func=cmd_save)

    # backup
    p = sub.add_parser('backup', help='Create .BACKUP of PBIX')
    p.add_argument('pbix'); p.set_defaults(func=cmd_backup)

    # measures
    p = sub.add_parser('measures', help='Manage measures via SSAS')
    me = p.add_subparsers(dest='measure_cmd', required=True)
    me.add_parser('list').set_defaults(func=cmd_list_measures)
    pa = me.add_parser('add')
    pa.set_defaults(func=cmd_add_measures)
    pa.add_argument('--measures-list', help='Comma-separated measure names')
    me.add_parser('fix-formats').set_defaults(func=cmd_fix_formats)
    me.add_parser('audit').set_defaults(func=cmd_audit_dax)
    pe = me.add_parser('export')
    pe.set_defaults(func=cmd_export_reference)
    pe.add_argument('-o', '--output')

    me.add_parser('inject-v16').set_defaults(func=cmd_inject_v16)

    args = parser.parse_args()
    args.func(args)

if __name__ == '__main__':
    main()
