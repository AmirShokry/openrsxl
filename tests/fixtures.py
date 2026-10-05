"""
Build minimal xlsx packages from raw XML snippets, so that each edge case can
be isolated in its own tiny fixture.
"""

import io
import zipfile

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

DEFAULT_STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="{MAIN}">
<numFmts count="3">
<numFmt numFmtId="164" formatCode="[h]:mm:ss"/>
<numFmt numFmtId="165" formatCode="yyyy-mm-dd"/>
<numFmt numFmtId="166" formatCode="0.000"/>
</numFmts>
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"><alignment horizontal="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

DEFAULT_STRINGS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="{MAIN}" count="4" uniqueCount="4"><si><t>zero</t></si><si><t>one</t></si><si><t xml:space="preserve"> two </t></si><si><r><t>th</t></r><r><rPr><b/></rPr><t>ree</t></r></si></sst>"""


def sheet_xml(sheet_data="", before="", after="", root_attrs="", decl=True):
    head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' if decl else ""
    return (
        f'{head}<worksheet xmlns="{MAIN}" xmlns:r="{REL}"{root_attrs}>'
        f"{before}<sheetData>{sheet_data}</sheetData>{after}</worksheet>"
    )


def make_xlsx(sheets=None, strings=DEFAULT_STRINGS, styles=DEFAULT_STYLES, date1904=False, raw_sheets=None):
    """
    sheets: list of sheetData snippets (str) or full sheet xml (if raw)
    raw_sheets: list of full sheet documents (str or bytes)
    """
    if raw_sheets is None:
        raw_sheets = [sheet_xml(s) for s in (sheets or [""])]
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
        overrides = [
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        ]
        if strings is not None:
            overrides.append(
                '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            )
        for i in range(len(raw_sheets)):
            overrides.append(
                f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            )
        z.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>' + "".join(overrides) + "</Types>",
        )
        z.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        )
        sheets_xml = "".join(
            f'<sheet name="Sheet{i+1}" sheetId="{i+1}" r:id="rId{i+1}"/>' for i in range(len(raw_sheets))
        )
        pr = '<workbookPr date1904="1"/>' if date1904 else ""
        z.writestr(
            "xl/workbook.xml",
            f'<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{REL}">{pr}<sheets>{sheets_xml}</sheets></workbook>',
        )
        rels = "".join(
            f'<Relationship Id="rId{i+1}" Type="{REL}/worksheet" Target="worksheets/sheet{i+1}.xml"/>'
            for i in range(len(raw_sheets))
        )
        n = len(raw_sheets)
        rels += f'<Relationship Id="rId{n+1}" Type="{REL}/styles" Target="styles.xml"/>'
        if strings is not None:
            rels += f'<Relationship Id="rId{n+2}" Type="{REL}/sharedStrings" Target="sharedStrings.xml"/>'
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>',
        )
        z.writestr("xl/styles.xml", styles)
        if strings is not None:
            z.writestr("xl/sharedStrings.xml", strings if isinstance(strings, bytes) else strings.encode("utf-8"))
        for i, s in enumerate(raw_sheets):
            z.writestr(f"xl/worksheets/sheet{i+1}.xml", s if isinstance(s, bytes) else s.encode("utf-8"))
    return b.getvalue()


def cells(*cs, r=1, row_attrs=""):
    """Wrap cell XML snippets in a row."""
    return f'<row r="{r}"{row_attrs}>' + "".join(cs) + "</row>"


def strings_xml(*sis, attrs=""):
    return f'<?xml version="1.0" encoding="UTF-8"?><sst xmlns="{MAIN}"{attrs}>' + "".join(sis) + "</sst>"
