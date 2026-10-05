"""
Reader edge cases: every fixture isolates one behaviour and is compared
against openpyxl (the oracle) in every load mode.
"""

import pytest
from fixtures import MAIN, cells, make_xlsx, sheet_xml, strings_xml
from oracle.compare import compare_workbooks

MODES = [
    {},
    {"read_only": True},
    {"data_only": True},
    {"rich_text": True},
    {"read_only": True, "data_only": True},
]


def c(value=None, t=None, s=None, r="A1", inner=None, extra=""):
    attrs = ""
    if r is not None:
        attrs += f' r="{r}"'
    if t is not None:
        attrs += f' t="{t}"'
    if s is not None:
        attrs += f' s="{s}"'
    attrs += extra
    if inner is None:
        inner = "" if value is None else f"<v>{value}</v>"
    return f"<c{attrs}>{inner}</c>"


NUMBERS = [
    "0",
    "1",
    "-1",
    "-0",
    "1.0",
    "1.5",
    "-1.5",
    "1E3",
    "1e-3",
    "1E+30",
    ".5",
    "5.",
    "1.7976931348623157e308",
    "1e309",
    "-1e309",
    "4.9e-324",
    "123456789012345678",
    "1234567890123456789012345678901234567890",
    "-9223372036854775808",
    "9223372036854775807",
    "9223372036854775808",
    "00012",
    "+7",
    "1_000",
    "1_0.5",
    " 12",
    "12 ",
    "\t3\n",
    "nan",
    "inf",
    "-inf",
    "Infinity",
    "1.2.3",
    "abc",
    "0x10",
    "١٢",
    "1e",
    "0.30000000000000004",
    "2.675",
    "1e-7",
    "1e16",
    "9007199254740993",
    "0.1E-2",
    "1.e5",
]


def number_cases():
    for n in NUMBERS:
        yield f"num-{n!r}", [cells(c(n))]
        # also through the date path
        yield f"date-{n!r}", [cells(c(n, s="1"))]


DATE_VALUES = [
    "0",
    "0.5",
    "0.999999",
    "0.9999999999",
    "1",
    "1.5",
    "59",
    "60",
    "60.5",
    "61",
    "-1",
    "-0.5",
    "2958465",
    "2958465.99999",
    "2958466",
    "1e20",
    "43832.12783564815",
    "44000.000005787",
    "1.00000578703704",
    "36526.999999999",
    "100000000",
]


def date_cases():
    for v in DATE_VALUES:
        for s in ("1", "2", "4", "5"):
            yield f"dateval-{v}-s{s}", [cells(c(v, s=s))]


STRING_CASES = {
    "ss-0": [cells(c("0", t="s"))],
    "ss-last": [cells(c("3", t="s"))],
    "ss-neg": [cells(c("-1", t="s"))],
    "ss-out-of-range": [cells(c("99", t="s"))],
    "ss-float-index": [cells(c("1.0", t="s"))],
    "ss-space-index": [cells(c(" 2 ", t="s"))],
    "ss-empty-v": [cells(c("", t="s"))],
    "bool-0": [cells(c("0", t="b"))],
    "bool-1": [cells(c("1", t="b"))],
    "bool-2": [cells(c("2", t="b"))],
    "bool-true": [cells(c("true", t="b"))],
    "err": [cells(c("#N/A", t="e"), c("#DIV/0!", t="e", r="B1"), c("whatever", t="e", r="C1"))],
    "str": [cells(c("hello", t="str"), c(" spaced ", t="str", r="B1"), c("a&amp;b&lt;c&gt;", t="str", r="C1"))],
    "str-crlf": [cells(c("a\r\nb\rc\nd", t="str"))],
    "str-charref": [cells(c("&#13;&#10;&#x41;&#233;&#x1F600;", t="str"))],
    "str-cdata": [cells(c(None, t="str", inner="<v><![CDATA[<x>&amp;]]></v>"))],
    "str-comment-in-v": [cells(c(None, t="str", inner="<v>ab<!-- c -->cd<?pi x?>ef</v>"))],
    "str-child-in-v": [cells(c(None, t="str", inner="<v>ab<x/>cd</v>"))],
    "iso-date": [
        cells(
            c("2020-01-02", t="d"),
            c("2020-01-02T03:04:05", t="d", r="B1"),
            c("12:30", t="d", r="C1"),
            c("2020-01-02T03:04:05.123Z", t="d", r="D1"),
            c("PT2H3M4S", t="d", r="E1"),
        )
    ],
    "iso-date-invalid": [cells(c("not a date", t="d"))],
    "unknown-type": [cells(c("x", t="zzz"))],
    "unknown-type-empty": [cells(c(None, t="zzz"))],
    "n-empty-v": [cells(c(None, inner="<v></v>"))],
    "n-ws-v": [cells(c(None, inner="<v> </v>"))],
    "two-v": [cells(c(None, inner="<v>1</v><v>2</v>"))],
    "v-other-ns": [cells(c(None, inner='<x:v xmlns:x="urn:other">1</x:v>'))],
    "inline": [cells(c(None, t="inlineStr", inner="<is><t>inline</t></is>"))],
    "inline-ws": [cells(c(None, t="inlineStr", inner='<is><t xml:space="preserve">  a  </t></is>'))],
    "inline-runs": [
        cells(c(None, t="inlineStr", inner='<is><r><t>a</t></r><r><rPr><b/><sz val="12"/></rPr><t>b</t></r></is>'))
    ],
    "inline-plain-and-runs": [cells(c(None, t="inlineStr", inner="<is><t>p</t><r><t>a</t></r></is>"))],
    "inline-empty": [cells(c(None, t="inlineStr", inner="<is/>"))],
    "inline-empty-t": [cells(c(None, t="inlineStr", inner="<is><t/></is>"))],
    "inline-missing-is": [cells(c(None, t="inlineStr"))],
    "inline-with-v": [cells(c(None, t="inlineStr", inner="<v>5</v><is><t>x</t></is>"))],
    "inline-phonetic": [
        cells(
            c(
                None,
                t="inlineStr",
                inner='<is><t>kanji</t><rPh sb="0" eb="1"><t>k</t></rPh><phoneticPr fontId="0" type="noConversion"/></is>',
            )
        )
    ],
    "inline-bad-rpr": [cells(c(None, t="inlineStr", inner='<is><r><rPr><sz val="abc"/></rPr><t>b</t></r></is>'))],
    "inline-unknown-child": [cells(c(None, t="inlineStr", inner="<is><foo/><t>x</t></is>"))],
    "inline-alias-child": [cells(c(None, t="inlineStr", inner="<is><plain>y</plain><t>x</t></is>"))],
    "inline-attr": [cells(c(None, t="inlineStr", inner='<is foo="1"><t>x</t></is>'))],
    "inline-two-t": [cells(c(None, t="inlineStr", inner="<is><t>x</t><t>y</t></is>"))],
    "inline-t-child": [cells(c(None, t="inlineStr", inner="<is><t>x<b>z</b>y</t></is>"))],
    "inline-r-empty": [cells(c(None, t="inlineStr", inner="<is><r/><r><t/></r></is>"))],
    "inline-x005F": [cells(c(None, t="inlineStr", inner="<is><t>a_x005F_x000D_b</t></is>"))],
    "inline-formula": [cells(c(None, t="inlineStr", inner="<f>1+1</f><is><t>x</t></is>"))],
}

STYLE_CASES = {
    "style-missing": [cells(c("1"))],
    "style-0": [cells(c("1", s="0"))],
    "style-3": [cells(c("1", s="3"))],
    "style-6": [cells(c("1.5", s="6"))],
    "style-empty": [cells(c("1", s=""))],
    "style-nonint": [cells(c("1", s="x"))],
    "style-out-of-range": [cells(c("1", s="99"))],
    "style-negative": [cells(c("1", s="-1"))],
    "style-space": [cells(c("1", s=" 3"))],
    "style-float": [cells(c("1", s="1.0"))],
}

COORD_CASES = {
    "coord-missing": [cells(c("1", r=None), c("2", r=None), c("3", r="E1"), c("4", r=None))],
    "coord-lower": [cells(c("1", r="b2"))],
    "coord-abs": [cells(c("1", r="$A$1"))],
    "coord-zero-pad": [cells(c("1", r="A01"))],
    "coord-row0": [cells(c("1", r="A0"))],
    "coord-max": [cells(c("1", r="XFD1048576"))],
    "coord-zzz": [cells(c("1", r="ZZZ5"))],
    "coord-too-long": [cells(c("1", r="AAAA1"))],
    "coord-empty": [cells(c("1", r=""), c("2", r=""))],
    "coord-no-digits": [cells(c("1", r="AB"))],
    "coord-duplicate": [cells(c("1", r="A1"), c("2", r="A1"))],
    "coord-other-row": [cells(c("1", r="C7"), r=2)],
    "coord-unordered": [cells(c("1", r="C1"), c("2", r="A1"))],
}

ROW_CASES = {
    "rows-no-r": ["<row><c><v>1</v></c></row><row><c><v>2</v></c></row>"],
    "rows-float-r": ['<row r="2.0"><c><v>1</v></c></row>'],
    "rows-bad-float-r": ['<row r="2.5"><c><v>1</v></c></row>'],
    "rows-space-r": ['<row r=" 3 "><c><v>1</v></c></row>'],
    "rows-unordered": ['<row r="5"><c r="A5"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>'],
    "rows-dims": [
        '<row r="1" ht="30" customHeight="1" hidden="1" outlineLevel="2" collapsed="1" s="3" customFormat="1"><c r="A1"><v>1</v></c></row>'
    ],
    "rows-spans-only": ['<row r="1" spans="1:3"><c r="A1"><v>1</v></c></row>'],
    "rows-dydescent": [
        '<row r="1" spans="1:3" xmlns:x14ac="urn:x14ac" x14ac:dyDescent="0.25"><c r="A1"><v>1</v></c></row>'
    ],
    "rows-dydescent-ht": [
        '<row r="1" ht="15" xmlns:x14ac="urn:x14ac" x14ac:dyDescent="0.25"><c r="A1"><v>1</v></c></row>'
    ],
    "rows-empty": ['<row r="1"/><row r="3" ht="20" customHeight="1"/>'],
    "rows-dup": ['<row r="1" ht="10"><c r="A1"><v>1</v></c></row><row r="1" ht="20"><c r="B1"><v>2</v></c></row>'],
    "rows-unknown-attr": ['<row r="1" foo="bar"><c r="A1"><v>1</v></c></row>'],
    "row-non-c-child": ['<row r="1"><x r="A1"><v>1</v></x><c r="B1"><v>2</v></c></row>'],
    "row-gap": ['<row r="1"><c r="A1"><v>1</v></c></row><row r="10"><c r="C10"><v>2</v></c></row>'],
    "row-many-cols": ['<row r="1">' + "".join(f'<c r="{chr(65 + i)}1"><v>{i}</v></c>' for i in range(26)) + "</row>"],
}

F = "SUM(A1:B2)"
FORMULA_CASES = {
    "f-plain": [cells(c(None, inner=f"<f>{F}</f><v>3</v>"))],
    "f-empty": [cells(c(None, inner="<f/><v>3</v>"))],
    "f-str-cached": [cells(c(None, t="str", inner="<f>A1&amp;B1</f><v>ab</v>"))],
    "f-err-cached": [cells(c(None, t="e", inner="<f>1/0</f><v>#DIV/0!</v>"))],
    "f-bool-cached": [cells(c(None, t="b", inner="<f>TRUE()</f><v>1</v>"))],
    "f-array": [cells(c(None, inner='<f t="array" ref="A1:A3">SUM(B1:B3*C1:C3)</f><v>1</v>'))],
    "f-array-noref": [cells(c(None, inner='<f t="array">1+1</f><v>2</v>'))],
    "f-datatable": [cells(c(None, inner='<f t="dataTable" ref="B2:C3" dt2D="1" dtr="1" r1="A1" r2="B1"/><v>1</v>'))],
    "f-datatable-extra-attr": [cells(c(None, inner='<f t="dataTable" ref="B2:C3" foo="1"/><v>1</v>'))],
    "f-unknown-type": [cells(c(None, inner='<f t="weird">1+1</f><v>2</v>'))],
    "f-shared": [
        '<row r="1">'
        '<c r="C1"><f t="shared" ref="C1:C4" si="0">A1+B1*$A$1+SUM(A:A)+SUM(1:1)+Sheet2!A1+\'My Sheet\'!B2</f><v>1</v></c>'
        '</row><row r="2"><c r="C2"><f t="shared" si="0"/><v>2</v></c><c r="D2"><f t="shared" si="0"/></c></row>'
        '<row r="3"><c r="E3"><f t="shared" si="0"/></c></row>'
    ],
    "f-shared-names-strings": [
        cells(
            c(
                None,
                r="B2",
                inner='<f t="shared" ref="B2:B3" si="1">IF(A1="A1",myName,"x"&amp;A2)+[1]Sheet1!A1+A1:A2:A3+#REF!+TRUE+1E+5+1.5E-3</f>',
            )
        ),
        cells(c(None, r="B3", inner='<f t="shared" si="1"/>'), r=3),
    ],
    "f-shared-out-of-range": [
        cells(c(None, r="B2", inner='<f t="shared" ref="A1:B2" si="0">A1</f>'), r=2),
        cells(c(None, r="A1", inner='<f t="shared" si="0"/>'), r=1),
    ],
    "f-shared-col-out": [
        cells(c(None, r="B1", inner='<f t="shared" si="0">A1</f>'), c(None, r="A1", inner='<f t="shared" si="0"/>'))
    ],
    "f-shared-no-master": [cells(c(None, inner='<f t="shared" si="7"/>'))],
    "f-shared-empty-master": [
        cells(
            c(None, r="A1", inner='<f t="shared" si="0"/>'),
            c(None, r="B1", inner='<f t="shared" si="0">A2</f>'),
            c(None, r="C1", inner='<f t="shared" si="0"/>'),
        )
    ],
    "f-shared-no-si": [cells(c(None, r="A1", inner='<f t="shared">B1</f>'), c(None, r="A2", inner='<f t="shared"/>'))],
    "f-shared-no-coord": [cells(c(None, r=None, inner='<f t="shared" si="0">B1</f>'))],
    "f-shared-dep-no-coord": [
        cells(c(None, r="A1", inner='<f t="shared" si="0">B1</f>'), c(None, r=None, inner='<f t="shared" si="0"/>'))
    ],
    "f-shared-lower-coord": [
        cells(c(None, r="a1", inner='<f t="shared" si="0">B1</f>'), c(None, r="a2", inner='<f t="shared" si="0"/>'))
    ],
    "f-shared-ws": [
        cells(
            c(None, r="A1", inner='<f t="shared" si="0">A1 + B1\n+C1</f>'),
            c(None, r="A2", inner='<f t="shared" si="0"/>'),
        )
    ],
    "f-shared-lower-ref": [
        cells(
            c(None, r="A1", inner='<f t="shared" si="0">b1+$c$1+c$1+$c1+zz9+xfd1048576</f>'),
            c(None, r="B3", inner='<f t="shared" si="0"/>'),
        )
    ],
    "f-shared-literal": [
        cells(c(None, r="A1", inner='<f t="shared" si="0">=A1</f>'), c(None, r="A2", inner='<f t="shared" si="0"/>'))
    ],
    "f-shared-bad-tokens": [cells(c(None, r="A1", inner='<f t="shared" si="0">(A1))</f>'))],
    "f-shared-bad-string": [cells(c(None, r="A1", inner='<f t="shared" si="0">"abc</f>'))],
    "f-shared-bad-bracket": [cells(c(None, r="A1", inner='<f t="shared" si="0">[1</f>'))],
    "f-shared-bad-error": [cells(c(None, r="A1", inner='<f t="shared" si="0">#FOO</f>'))],
    "f-shared-mismatch": [cells(c(None, r="A1", inner='<f t="shared" si="0">{1,2)</f>'))],
    "f-shared-unicode": [
        cells(
            c(None, r="A1", inner='<f t="shared" si="0">A1+١+"é"</f>'), c(None, r="A2", inner='<f t="shared" si="0"/>')
        )
    ],
    "f-shared-array-literal": [
        cells(
            c(None, r="A1", inner='<f t="shared" si="0">SUM({1,2;3,4})+A1</f>'),
            c(None, r="B2", inner='<f t="shared" si="0"/>'),
        )
    ],
    "f-shared-sci": [
        cells(
            c(None, r="A1", inner='<f t="shared" si="0">1E+3+A1-2e-1</f>'),
            c(None, r="A2", inner='<f t="shared" si="0"/>'),
        )
    ],
    "f-shared-ops": [
        cells(
            c(None, r="A1", inner='<f t="shared" si="0">-A1%^2&gt;=B1&lt;&gt;C1&amp;D1</f>'),
            c(None, r="A2", inner='<f t="shared" si="0"/>'),
        )
    ],
    "f-shared-dependent-text": [
        cells(
            c(None, r="A1", inner='<f t="shared" si="0">B1</f>'), c(None, r="A2", inner='<f t="shared" si="0">ZZ99</f>')
        )
    ],
    "f-other-ns": [cells(c(None, inner='<f xmlns="urn:x">1+1</f><v>2</v>'))],
    "f-with-children": [cells(c(None, inner="<f>A1<x/>B1</f>"))],
}

XML_CASES = {
    "xml-prefixed": [
        f'<?xml version="1.0"?><x:worksheet xmlns:x="{MAIN}"><x:sheetData><x:row r="1"><x:c r="A1" x:t="s"><x:v>1</x:v></x:c><x:c r="B1" t="s"><x:v>2</x:v></x:c></x:row></x:sheetData></x:worksheet>'
    ],
    "xml-no-decl": [sheet_xml(cells(c("1")), decl=False)],
    "xml-bom": ["﻿" + sheet_xml(cells(c("1")))],
    "xml-crlf-attrs": [sheet_xml('<row r="1"><c r="A1"\r\n t="str"><v>a</v></c></row>')],
    "xml-attr-ws": [
        sheet_xml('<row r="1"><c r="A1" t="e"><v>x</v></c></row>', before='<sheetPr codeName="a&#9;b\tc\nd"/>')
    ],
    "xml-single-quotes": [sheet_xml("<row r='1'><c r='A1' t='str'><v>q</v></c></row>")],
    "xml-comments-pis": [sheet_xml('<!-- c --><row r="1"><?pi?><c r="A1"><!--x--><v>1</v></c></row>')],
    "xml-doctype": [
        '<?xml version="1.0"?><!DOCTYPE worksheet><worksheet xmlns="'
        + MAIN
        + '"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'
    ],
    "xml-latin1": [
        ('<?xml version="1.0" encoding="ISO-8859-1"?>' + sheet_xml(cells(c("café", t="str")), decl=False)).encode(
            "latin-1"
        )
    ],
    "xml-utf16": [sheet_xml(cells(c("x", t="str"))).replace('encoding="UTF-8"', 'encoding="UTF-16"').encode("utf-16")],
    "xml-malformed-unclosed": [
        '<?xml version="1.0"?><worksheet xmlns="'
        + MAIN
        + '"><sheetData><row r="1"><c r="A1"><v>1</v></row></sheetData></worksheet>'
    ],
    "xml-malformed-entity": [sheet_xml(cells(c("&nbsp;", t="str")))],
    "xml-malformed-amp": [sheet_xml(cells(c("a & b", t="str")))],
    "xml-malformed-lt-attr": [sheet_xml('<row r="1"><c r="A<1"><v>1</v></c></row>')],
    "xml-malformed-dup-attr": [sheet_xml('<row r="1"><c r="A1" r="B1"><v>1</v></c></row>')],
    "xml-malformed-ctrl": [sheet_xml(cells(c("a\x01b", t="str")))],
    "xml-malformed-junk": [sheet_xml(cells(c("1"))) + "<extra/>"],
    "xml-malformed-text-after": [sheet_xml(cells(c("1"))) + "junk"],
    "xml-malformed-unbound": [sheet_xml(cells(c(None, inner="<y:v>1</y:v>")))],
    "xml-malformed-noattrspace": [sheet_xml('<row r="1"><c r="A1"t="s"><v>1</v></c></row>')],
    "xml-malformed-charref": [sheet_xml(cells(c("&#0;", t="str")))],
    "xml-malformed-cdata-end": [sheet_xml(cells(c("a]]>b", t="str")))],
    "xml-empty": [""],
    "xml-strict-ns": [
        '<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'
    ],
    "xml-other-root": [f'<foo xmlns="{MAIN}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></foo>'],
    "xml-row-outside": [sheet_xml(cells(c("1")), after='<extLst><row r="9"><c r="A9"><v>9</v></c></row></extLst>')],
    "xml-sheetdata-junk": [sheet_xml(cells(c("1")) + '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>')],
    "xml-extlst-in-cell": [sheet_xml(cells(c(None, inner="<v>1</v><extLst/>")))],
    "xml-col-in-cell": [sheet_xml(cells(c(None, inner='<v>1</v><col min="1" max="1" width="5"/>')))],
    "xml-self-closing-sheetdata": [f'<worksheet xmlns="{MAIN}"><sheetData/></worksheet>'],
    "xml-no-sheetdata": [
        f'<worksheet xmlns="{MAIN}"><sheetViews><sheetView workbookViewId="0"/></sheetViews></worksheet>'
    ],
    "xml-big-text": [sheet_xml(cells(c("x" * 40000, t="str"), c("&amp;" * 20000, t="str", r="B1")))],
}

ELEMENT_CASES = {
    "el-cols": [
        sheet_xml(
            cells(c("1")),
            before='<cols><col min="1" max="3" width="12.5" customWidth="1" style="3"/><col min="5" max="5" hidden="1"/></cols>',
        )
    ],
    "el-merge": [
        sheet_xml(cells(c("1"), c("2", r="B1")), after='<mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>')
    ],
    "el-merge-bad": [sheet_xml(cells(c("1")), after='<mergeCells><mergeCell ref="garbage"/></mergeCells>')],
    "el-cf": [
        sheet_xml(
            cells(c("1")),
            after='<conditionalFormatting sqref="A1:A5"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>',
        )
    ],
    "el-cf-bad": [
        sheet_xml(
            cells(c("1")),
            after='<conditionalFormatting sqref="A1"><cfRule type="nope" priority="x"/></conditionalFormatting>',
        )
    ],
    "el-dv": [
        sheet_xml(
            cells(c("1")),
            after='<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="A1:A3"><formula1>"a,b"</formula1></dataValidation></dataValidations>',
        )
    ],
    "el-extlst": [
        sheet_xml(
            cells(c("1")),
            after='<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"/><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"/></extLst>',
        )
    ],
    "el-sheetpr": [
        sheet_xml(
            cells(c("1")),
            before='<sheetPr codeName="Sheet1" filterMode="0"><tabColor rgb="FFFF0000"/><outlinePr summaryBelow="0"/><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:C3"/><sheetViews><sheetView tabSelected="1" workbookViewId="0" zoomScale="85"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="B2" sqref="B2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25" xmlns:x14ac="urn:x14ac"/>',
        )
    ],
    "el-print": [
        sheet_xml(
            cells(c("1")),
            after='<printOptions gridLines="1"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup orientation="landscape" paperSize="9"/><headerFooter><oddHeader>&amp;CTitle</oddHeader></headerFooter><rowBreaks count="1" manualBreakCount="1"><brk id="5" max="16383" man="1"/></rowBreaks><colBreaks count="1"><brk id="3" man="1"/></colBreaks>',
        )
    ],
    "el-protection": [
        sheet_xml(cells(c("1")), after='<sheetProtection password="CC1A" sheet="1" objects="1" scenarios="1"/>')
    ],
    "el-autofilter": [
        sheet_xml(
            cells(c("1")),
            after='<autoFilter ref="A1:C5"><filterColumn colId="0"><filters><filter val="1"/></filters></filterColumn></autoFilter>',
        )
    ],
    "el-custom-views": [
        sheet_xml(
            cells(c("1")),
            after='<rowBreaks count="1"><brk id="5" man="1"/></rowBreaks><customSheetViews><customSheetView guid="{00000000-0000-0000-0000-000000000000}"/></customSheetViews>',
        )
    ],
    "el-hyperlink-internal": [
        sheet_xml(
            cells(c("1"), c("2", r="B1")),
            after='<hyperlinks><hyperlink ref="A1" location="Sheet1!B1" display="go"/><hyperlink ref="B1:C2" location="X"/></hyperlinks>',
        )
    ],
    "el-dimension-bad": [sheet_xml(cells(c("1")), before='<dimension ref="garbage"/>')],
    "el-dimension-after": [sheet_xml(cells(c("1"))).replace("</worksheet>", '<dimension ref="A1:Z9"/></worksheet>')],
    "el-dimension-extra-attr": [sheet_xml(cells(c("1")), before='<dimension ref="A1:B2" foo="1"/>')],
    "el-unknown": [sheet_xml(cells(c("1")), before="<foo><bar/></foo>", after="<baz/>")],
    "el-scenarios": [
        sheet_xml(
            cells(c("1")),
            after='<scenarios current="0"><scenario name="s" locked="1" count="1"><inputCells r="A1" val="2"/></scenario></scenarios>',
        )
    ],
    "el-legacy": [sheet_xml(cells(c("1")), after='<legacyDrawing r:id="rId1"/>')],
}

SHARED_STRING_CASES = {
    "sst-plain": strings_xml("<si><t>a</t></si>", "<si><t>b</t></si>"),
    "sst-empty-si": strings_xml("<si/>", "<si></si>", "<si><t/></si>"),
    "sst-ws": strings_xml('<si><t xml:space="preserve">  a  </t></si>', "<si><t>\n</t></si>"),
    "sst-runs": strings_xml(
        '<si><r><t>a</t></r><r><rPr><b/><i/><sz val="11"/><color rgb="FFFF0000"/><rFont val="Arial"/><family val="2"/><scheme val="minor"/></rPr><t>b</t></r></si>'
    ),
    "sst-x005F": strings_xml("<si><t>_x005F_x000D_ and x005F_ and _x000D_</t></si>"),
    "sst-phonetic": strings_xml('<si><t>a</t><rPh sb="0" eb="1"><t>b</t></rPh><phoneticPr fontId="1"/></si>'),
    "sst-bad-rpr": strings_xml('<si><r><rPr><u val="nope"/></rPr><t>x</t></r></si>'),
    "sst-attr": strings_xml('<si foo="1"><t>x</t></si>'),
    "sst-nested-si": strings_xml("<si><t>x</t><si><t>y</t></si></si>"),
    "sst-other-elements": strings_xml('<foo/><si><t>x</t></si><extLst><ext uri="x"/></extLst>'),
    "sst-entities": strings_xml("<si><t>&lt;&amp;&gt;&quot;&apos;&#9;&#xD;</t></si>"),
    "sst-crlf": strings_xml("<si><t>a\r\nb\rc</t></si>"),
    "sst-cdata": strings_xml("<si><t><![CDATA[<b>]]></t></si>"),
    "sst-prefixed": f'<x:sst xmlns:x="{MAIN}"><x:si><x:t>p</x:t></x:si></x:sst>',
    "sst-malformed": strings_xml("<si><t>x</t>"),
    "sst-doctype": '<?xml version="1.0"?><!DOCTYPE sst>' + strings_xml("<si><t>x</t></si>"),
    "sst-many": strings_xml(*[f"<si><t>s{i}</t></si>" for i in range(5000)]),
}


def all_cases():
    yield from number_cases()
    yield from date_cases()
    for group in (STRING_CASES, STYLE_CASES, COORD_CASES, ROW_CASES, FORMULA_CASES):
        for k, v in group.items():
            yield k, v
    for k, v in XML_CASES.items():
        yield k, ("raw", v)
    for k, v in ELEMENT_CASES.items():
        yield k, ("raw", v)


CASES = dict(all_cases())


def build(case):
    if isinstance(case, tuple) and case[0] == "raw":
        return make_xlsx(raw_sheets=case[1])
    return make_xlsx(case)


@pytest.mark.parametrize("mode", MODES, ids=lambda m: "-".join(m) or "default")
@pytest.mark.parametrize("name", list(CASES))
def test_sheet_case(name, mode):
    data = build(CASES[name])
    diffs = compare_workbooks(data, **mode)
    assert not diffs, "\n".join(diffs)


@pytest.mark.parametrize("mode", MODES, ids=lambda m: "-".join(m) or "default")
@pytest.mark.parametrize("name", list(SHARED_STRING_CASES))
def test_shared_strings(name, mode):
    n = 3
    rows = cells(*[c(str(i), t="s", r=f"A{i+1}") for i in range(n)])
    data = make_xlsx([rows], strings=SHARED_STRING_CASES[name])
    diffs = compare_workbooks(data, **mode)
    assert not diffs, "\n".join(diffs)


@pytest.mark.parametrize("mode", MODES, ids=lambda m: "-".join(m) or "default")
def test_date1904(mode):
    data = make_xlsx([cells(*[c(v, s="1", r=f"A{i+1}") for i, v in enumerate(DATE_VALUES)])], date1904=True)
    assert not compare_workbooks(data, **mode)


@pytest.mark.parametrize("mode", MODES, ids=lambda m: "-".join(m) or "default")
def test_multiple_sheets(mode):
    data = make_xlsx([cells(c("1")), cells(c("0", t="s")), ""])
    assert not compare_workbooks(data, **mode)


# ElementTree.iterparse feeds expat 16KB at a time: a well-formedness error
# hides every event completed in the same block, so whether openpyxl reports
# a bad cell value or the syntax error depends on block boundaries.
def _block_case(pad_cells, gap_cells, bad="&lt;bad", error="<<x"):
    pad = "".join(f'<c r="A{i}"><v>{i}</v></c>' for i in range(1, pad_cells + 1))
    gap = "".join(f'<c r="B{i}"><v>{i}</v></c>' for i in range(1, gap_cells + 1))
    rows = (
        f'<row r="1">{pad}</row><row r="2"><c r="{bad}1"><v>1</v></c></row>'
        f'<row r="3">{gap}</row>{error}<row r="4"><c r="A4"><v>4</v></c></row>'
    )
    return make_xlsx(raw_sheets=[sheet_xml(rows)])


@pytest.mark.parametrize("mode", [{}, {"read_only": True}], ids=["default", "read_only"])
@pytest.mark.parametrize("pad", [0, 500, 560, 580, 600, 620, 1200, 1210, 1220])
@pytest.mark.parametrize("gap", [0, 5, 50, 700])
def test_feed_block_semantics(pad, gap, mode):
    assert not compare_workbooks(_block_case(pad, gap), check_save=False, **mode)


# regression: more cells with a non-standard data type than codes of one byte
# (the type of every such cell used to get its own code, which overflowed)
MANY_TYPES = {
    "str_without_value": "".join(
        f'<row r="{r}"><c r="A{r}" t="str"><f>"x"</f></c><c r="B{r}" t="str"><v>v</v></c></row>' for r in range(1, 400)
    ),
    "many_distinct_types": "".join(
        f'<row r="{r}"><c r="A{r}" t="t{r}"><v>{r}</v></c><c r="B{r}" t="str"/></row>' for r in range(1, 300)
    ),
}


@pytest.mark.parametrize("mode", MODES, ids=lambda m: "-".join(m) or "default")
@pytest.mark.parametrize("name", list(MANY_TYPES))
def test_many_extra_data_types(name, mode):
    data = make_xlsx([MANY_TYPES[name]])
    diffs = compare_workbooks(data, **mode)
    assert not diffs, "\n".join(diffs)
