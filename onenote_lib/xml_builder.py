"""Build valid OneNote XML fragments for writing content.

Uses string construction (not xml.etree.ElementTree) because ET cannot
produce CDATA sections, which OneNote requires for text content.
"""

ONENOTE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


def cdata(text: str) -> str:
    """Wrap text in CDATA, escaping any ]]> sequences."""
    return f"<![CDATA[{text.replace(']]>', ']]]]><![CDATA[>')}]]>"


def build_text_oe(text: str, bold: bool = False, italic: bool = False) -> str:
    """Build an OE element with text content and optional formatting."""
    styled = text
    if bold and italic:
        styled = f'<span style="font-weight:bold;font-style:italic">{styled}</span>'
    elif bold:
        styled = f'<span style="font-weight:bold">{styled}</span>'
    elif italic:
        styled = f'<span style="font-style:italic">{styled}</span>'
    return f"<one:OE><one:T>{cdata(styled)}</one:T></one:OE>"


def build_heading(text: str, level: int = 2) -> str:
    """Build a heading OE with appropriate font size."""
    sizes = {1: 20, 2: 16, 3: 13, 4: 12}
    size = sizes.get(level, 12)
    styled = f'<span style="font-size:{size}pt;font-weight:bold">{text}</span>'
    return f"<one:OE><one:T>{cdata(styled)}</one:T></one:OE>"


def build_list_item(text: str, numbered: bool = False) -> str:
    """Build a list item OE with bullet or number marker."""
    if numbered:
        list_elem = '<one:Number numberSequence="1" numberFormat="##." />'
    else:
        list_elem = '<one:Bullet bullet="2" fontSize="11.0" />'
    return (
        f"<one:OE>"
        f"<one:List>{list_elem}</one:List>"
        f"<one:T>{cdata(text)}</one:T>"
        f"</one:OE>"
    )


def build_table(rows: list[list[str]], has_header: bool = True) -> str:
    """Build a Table element from rows of cell strings.

    Structure: Table > Columns + Row > Cell > OEChildren > OE > T
    """
    if not rows:
        return ""
    col_count = max(len(r) for r in rows)
    cols_xml = "".join(
        f'<one:Column index="{i}" width="120" />' for i in range(col_count)
    )
    rows_xml = []
    for row in rows:
        cells = []
        for cell_text in row:
            cells.append(
                f"<one:Cell>"
                f"<one:OEChildren>"
                f"<one:OE><one:T>{cdata(cell_text)}</one:T></one:OE>"
                f"</one:OEChildren>"
                f"</one:Cell>"
            )
        while len(cells) < col_count:
            cells.append(
                f"<one:Cell>"
                f"<one:OEChildren>"
                f"<one:OE><one:T>{cdata('')}</one:T></one:OE>"
                f"</one:OEChildren>"
                f"</one:Cell>"
            )
        rows_xml.append(f"<one:Row>{''.join(cells)}</one:Row>")
    header_attr = "true" if has_header and len(rows) > 1 else "false"
    return (
        f'<one:Table bordersVisible="true" hasHeaderRow="{header_attr}">'
        f"<one:Columns>{cols_xml}</one:Columns>"
        f"{''.join(rows_xml)}"
        f"</one:Table>"
    )


def build_outline(oe_elements: list[str]) -> str:
    """Wrap OE element strings in an Outline > OEChildren structure."""
    return (
        f"<one:Outline>"
        f"<one:OEChildren>"
        f"{''.join(oe_elements)}"
        f"</one:OEChildren>"
        f"</one:Outline>"
    )
