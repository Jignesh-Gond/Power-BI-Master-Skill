---
name: power-bi-master
description: POWER BI MASTER SKILL by JG — The ultimate Power BI Automation and DAX Engineering standard. Master PBIX surgery, Lakh/Crore formatting, and high-performance DAX patterns. Compatible with Gemini CLI, OpenCode, and Agy CLI.
---

# 🚀 POWER BI MASTER SKILL by JG: Automation & Engineering Standard

> **Universal Engineering Standard for Professional Power BI Development.**
> This skill enables AI Agents (Gemini, OpenCode, Agy) to perform expert-level dashboard engineering, model auditing, and PBIX automation.

---

## 🌏 Tailored for Global Use, Optimized for India
Includes specialized logic for the **Indian Numbering System**, ensuring data is presented in **Lakhs (L)** and **Crores (Cr)** instead of Millions/Billions.

---

## 🛠 CRITICAL RULES (NEVER VIOLATE)

### File & ZIP Safety
- **Never modify PBIX while open**: Modifying the ZIP while Power BI Desktop is open leads to corruption.
- **DataModel Compression**: Always repack `DataModel` as `ZIP_STORED` (No compression). Everything else uses `ZIP_DEFLATED`.
- **Security Cleanup**: ALWAYS remove `SecurityBindings` on every repack to ensure portability.

### DAX & Model Integrity
- **MANDATORY DAX AUDIT**: After every rebuild, run a DAX audit to fix corruption (e.g., duplicated names like `Measure = Measure = VAR...`).
- **TOM Persistence**: `model.SaveChanges()` only saves to the live model; use `SendKeys` (Ctrl+S) or CLI tools to persist to the `.pbix` file.

### Visual JSON Standards
- **Subtitle Fix (CRITICAL)**: Force `subTitle.show = false` in 3 locations: `vco.subTitle` (Capital 'T'), `objs.subTitle` (Capital 'T'), and `objs.general.autoSubtitle`.
- **Coloring**: Avoid `ThemeDataColor` for text; use explicit `#333333` (DARK_TEXT) for readability.
- **Labels**: Charts use the `labels` key, NOT `dataLabels`.

---

## 📐 Visual Build Standards (CEO Style)

### Layout Reference (720px Canvas)
| Row | Visual Type | Y Pos | Height | Notes |
|-----|-------------|-------|--------|-------|
| 1 | KPI Cards | 5 | 80 | 4 cards, ~305px each |
| 2 | Slicers | 95 | 50 | 3-column layout (414px each) |
| 3 | Bar/Column | 160 | 260 | Descending sort by default |
| 4 | Bottom Trends| 435 | 250 | Line charts for trends |

### KPI Card Standards
- **Title**: Top-aligned, `#E6F3FF` background, 14pt font, centered.
- **Value**: Centered, `#17614A` (Deep Green), 22-24pt font.
- **Padding**: Set to `0D` to eliminate wasted space.

### Slicer Standards
- **Mode**: Dropdown preferred.
- **Selection**: "Select All" enabled, "Single Select" disabled.
- **Hierarchy**: Use `header.text` for field labels; hide `VCO.title`.

---

## 🧮 Universal DAX Patterns

### 1. Indian Custom Formatting (Dynamic)
Correctly places commas for Lakh/Crore in any visual.
```dax
Indian Format Display =
VAR Val = [Target Measure]
VAR SafeLog = IF(Val > 0, FLOOR(LOG10(Val + 0.0000001), 1), 0)
VAR GroupsAfter3 = MIN(3, MAX(0, ROUNDUP((SafeLog - 2) / 2, 0)))
VAR Pattern = IF(GroupsAfter3 = 0, "#,##0", REPT("#,##\\,", GroupsAfter3) & "##0")
RETURN IF(Val = 0 || ISBLANK(Val), "0", FORMAT(Val, Pattern))
```

### 2. Amount Display (Cr/Lakh Suffix)
```dax
Display Amount =
VAR v = [Raw Amount]
RETURN
SWITCH(TRUE(),
    v >= 10000000, FORMAT(v/10000000, "0.00") & " Cr",
    v >= 100000, FORMAT(v/100000, "0.00") & " L",
    FORMAT(v, "#,0")
)
```

### 3. Universal Growth Intelligence
```dax
Growth % =
VAR Prev = CALCULATE([Metric], DATEADD('Date'[Date], -1, MONTH))
RETURN DIVIDE([Metric] - Prev, Prev)
```

---

## 🤖 AI CLI & Automation Integration (Unified CLI)

This skill is powered by a high-performance, 2000-line automation suite: **`powerbi_automation.py`**. It performs direct surgery on the PBIX structure and live model injection.

### Automation Commands
| Category | Command | Description |
|----------|---------|-------------|
| **Surgery** | `fix-visuals` | Applies subtitle fixes, KPI styling, and display units. |
| **Surgery** | `build-v16` | Complete overhaul: auto-subtitles, desc sort, sidebar, executive summary. |
| **Surgery** | `rebuild-pbix` | Rebuilds report with a standard 4-page layout. |
| **Model** | `measures inject-v16`| Injects the universal DAX library (Indian units, growth, etc.) via SSAS. |
| **Model** | `measures audit` | Scans for DAX corruption and auto-fixes duplicated name errors. |
| **Analysis**| `analyze` | Detailed PBIX structure mapping (pages, visuals, fields). |
| **Utility** | `backup` | Creates a timestamped `.BACKUP` of the PBIX. |

### Professional Automation Workflow
1. **Extraction**: `extract_pbix(src, tmp_dir)` — Unpacks the PBIX ZIP.
2. **Security Cleanup**: `remove_security_bindings(tmp_dir)` — Ensures the file opens anywhere.
3. **JSON Patching**: Direct manipulation of `visual.json` and `pages.json` for pixel-perfect layouts.
4. **Repackaging**: `repack_pbix(tmp_dir, output)` — **Crucial**: Stores `DataModel` as `ZIP_STORED`.

### Key Automation Helpers (Internal)
```python
# Field Reference Helpers
def col_field(ent, prop): return {"Column": {"Expression": {"SourceRef": {"Entity": ent}}, "Property": prop}}
def meas_field(ent, prop): return {"Measure": {"Expression": {"SourceRef": {"Entity": ent}}, "Property": prop}}

# Subtitle Force-Disable (3-Level Fix)
def fix_subtitles(v_json):
    v_json['objects']['subTitle'] = [{"properties": {"show": L("false")}}]
    v_json['visualContainerObjects']['subTitle'] = [{"properties": {"show": L("false")}} ]
    v_json['objects'].setdefault('general', [{}])[0]['properties']['autoSubtitle'] = L("false")
```

---

## 📂 Universal Workspace Structure

```text
project-root/
├── src/
│   ├── main_report.pbix       # Core report file
│   └── reference_measures.dax # Exported DAX library
├── powerbi_automation.py      # Unified Automation CLI (v16+)
├── .opencode/
│   └── skills/
│       └── power-bi-master/  # POWER BI MASTER SKILL by JG
└── README.md                 # Project Overview
```

---

## ✅ Deployment Standards
1. **Analyze**: Use `pbi-cli-tool` to map model tables and measures.
2. **Apply Standards**: Run automation scripts to fix visual JSON (subtitles, colors, KPI styles).
3. **Audit**: Run `measures audit` to fix DAX corruption and inject universal Indian units.
4. **Persist**: Ensure `SecurityBindings` are removed before uploading to GitHub.

---

*Crafted by JG for the global Power BI community.*
**Give us a star on GitHub if you find this standard helpful!** ⭐
