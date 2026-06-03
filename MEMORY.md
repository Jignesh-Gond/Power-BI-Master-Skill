# 🧠 Power BI Master: Technical Memory

## Unified Automation Engine
The `powerbi_automation.py` script is the heart of this skill. It replaces 52 legacy scripts with a single, maintainable CLI.

### Key Logic: PBIX Surgery
1. **Subtitle Key Name**: Must be `subTitle` (Capital T). Lowercase `subtitle` is ignored by Power BI Desktop.
2. **Theme Patching**: The Fluent2 theme's `autoSubtitle` property must be patched globally in the theme JSON to prevent overrides.
3. **DataModel Storage**: Always use `zipfile.ZIP_STORED` for the `DataModel` file. Using compression will often corrupt the file or prevent Power BI from opening it.
4. **Security Bindings**: These are environment-specific. Deleting them and removing their references in `[Content_Types].xml` is mandatory for portable reports.

### SSAS (TOM) Connection
The script uses `pythonnet` to connect to the live Analysis Services instance inside an open Power BI Desktop window.
- **Port Discovery**: Scans local ports to find the dynamic SSAS port.
- **Audit logic**: Identifies measures where the name is repeated in the expression (a common export/import artifact) and cleans them.

### Indian Numbering Logic
The `Indian Format Display` DAX uses `LOG10` and `REPT` to dynamically place commas at Lakh (1,00,000) and Crore (1,00,00,000) intervals, which standard Power BI format strings cannot handle perfectly without localized OS settings.
