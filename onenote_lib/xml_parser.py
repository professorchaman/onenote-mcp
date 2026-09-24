"""Parse OneNote XML into markdown and structured data."""

import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

# OneNote 2013 XML namespace
NS = {"one": "http://schemas.microsoft.com/office/onenote/2013/onenote"}


@dataclass
class ImageRef:
    """Reference to an image in a OneNote page."""
    callback_id: str
    index: int
    width: float | None = None
    height: float | None = None
    alt_text: str | None = None


@dataclass
class PageInfo:
    """Parsed page metadata."""
    id: str
    name: str
    last_modified: str | None = None
    level: int = 0


@dataclass
class SectionInfo:
    """Parsed section metadata."""
    id: str
    name: str
    path: str | None = None
    pages: list[PageInfo] = field(default_factory=list)


@dataclass
class SectionGroupInfo:
    """Parsed section group metadata."""
    id: str
    name: str
    sections: list[SectionInfo] = field(default_factory=list)
    section_groups: list["SectionGroupInfo"] = field(default_factory=list)


@dataclass
class NotebookInfo:
    """Parsed notebook metadata."""
    id: str
    name: str
    path: str | None = None
    last_modified: str | None = None
    sections: list[SectionInfo] = field(default_factory=list)
    section_groups: list[SectionGroupInfo] = field(default_factory=list)


def parse_notebooks(xml_str: str) -> list[NotebookInfo]:
    """Parse hierarchy XML into notebook list."""
    root = ET.fromstring(xml_str)
    notebooks = []
    for nb in root.findall("one:Notebook", NS):
        notebook = NotebookInfo(
            id=nb.get("ID", ""),
            name=nb.get("name", ""),
            path=nb.get("path"),
            last_modified=nb.get("lastModifiedTime"),
        )
        notebook.sections = _parse_sections(nb)
        notebook.section_groups = _parse_section_groups(nb)
        notebooks.append(notebook)
    return notebooks


def _parse_sections(parent) -> list[SectionInfo]:
    """Parse Section elements under a parent node."""
    sections = []
    for sec in parent.findall("one:Section", NS):
        section = SectionInfo(
            id=sec.get("ID", ""),
            name=sec.get("name", ""),
            path=sec.get("path"),
        )
        for page in sec.findall("one:Page", NS):
            section.pages.append(PageInfo(
                id=page.get("ID", ""),
                name=page.get("name", ""),
                last_modified=page.get("lastModifiedTime"),
                level=int(page.get("pageLevel", "0")),
            ))
        sections.append(section)
    return sections


def _parse_section_groups(parent) -> list[SectionGroupInfo]:
    """Parse SectionGroup elements recursively."""
    groups = []
    for sg in parent.findall("one:SectionGroup", NS):
        # Skip recycle bin
        if sg.get("isRecycleBin") == "true":
            continue
        group = SectionGroupInfo(
            id=sg.get("ID", ""),
            name=sg.get("name", ""),
        )
        group.sections = _parse_sections(sg)
        group.section_groups = _parse_section_groups(sg)
        groups.append(group)
    return groups


def parse_page_to_markdown(xml_str: str) -> tuple[str, list[ImageRef]]:
    """Convert OneNote page XML to markdown text + image references.

    Returns:
        Tuple of (markdown_text, list_of_image_refs)
    """
    root = ET.fromstring(xml_str)
    title = root.get("name", root.get("ID", "Untitled"))
    lines = [f"# {title}", ""]

    tag_defs = _parse_tag_defs(root)

    images: list[ImageRef] = []
    img_counter = 0

    # Process all Outline elements (main content containers)
    for outline in root.findall(".//one:Outline", NS):
        outline_lines, outline_images, img_counter = _process_outline(
            outline, images_start_index=img_counter, tag_defs=tag_defs
        )
        lines.extend(outline_lines)
        images.extend(outline_images)
        lines.append("")

    # Process top-level images (outside outlines)
    for img in root.findall(".//one:Image", NS):
        # Skip images already found inside outlines
        cb_id = _get_callback_id(img)
        if cb_id and not any(i.callback_id == cb_id for i in images):
            img_counter += 1
            ref = _make_image_ref(img, img_counter)
            if ref:
                images.append(ref)
                lines.append(f"[Image {ref.index}]")

    return "\n".join(lines).strip(), images


def _process_outline(outline, images_start_index: int = 0, tag_defs: dict | None = None) -> tuple[list[str], list[ImageRef], int]:
    """Process an Outline element into markdown lines."""
    oe_children = outline.find("one:OEChildren", NS)
    if oe_children is None:
        return [], [], images_start_index
    lines, images, img_counter = _walk_oe_children(oe_children, images_start_index, tag_defs=tag_defs or {})
    return lines, images, img_counter


def _walk_oe_children(
    oe_children, img_counter: int, depth: int = 0, tag_defs: dict | None = None,
) -> tuple[list[str], list[ImageRef], int]:
    """Walk OE children with controlled dispatch — no double-visiting."""
    lines: list[str] = []
    images: list[ImageRef] = []
    if tag_defs is None:
        tag_defs = {}

    for oe in oe_children.findall("one:OE", NS):
        table = oe.find("one:Table", NS)
        if table is not None:
            lines.extend(_process_table(table))
            continue

        image = oe.find("one:Image", NS)
        if image is not None:
            cb_id = _get_callback_id(image)
            if cb_id:
                img_counter += 1
                ref = _make_image_ref(image, img_counter)
                if ref:
                    images.append(ref)
                    lines.append(f"[Image {ref.index}]")
            continue

        inserted = oe.find("one:InsertedFile", NS)
        if inserted is not None:
            name = inserted.get("preferredName", "file")
            lines.append(f"[Attached: {name}]")
            continue

        prefix = _get_tag_prefix(oe, tag_defs)
        list_prefix = _get_list_prefix(oe, depth)

        for t in oe.findall("one:T", NS):
            text = _clean_text(t.text or "")
            if text.strip():
                indent = "  " * depth
                line = f"{indent}{list_prefix}{prefix}{text}" if (depth > 0 or list_prefix or prefix) else text
                lines.append(line)

        nested = oe.find("one:OEChildren", NS)
        if nested is not None:
            sub_lines, sub_images, img_counter = _walk_oe_children(
                nested, img_counter, depth + 1, tag_defs=tag_defs,
            )
            lines.extend(sub_lines)
            images.extend(sub_images)

    return lines, images, img_counter


def _parse_tag_defs(root) -> dict:
    """Extract TagDef definitions from a page root. Returns {index: {name, type, completed}}."""
    defs = {}
    for td in root.findall("one:TagDef", NS):
        index = td.get("index", "")
        tag_type = td.get("type", "")
        name = td.get("name", "")
        defs[index] = {"name": name, "type": tag_type}
    return defs


def _get_tag_prefix(oe, tag_defs: dict) -> str:
    """Check if an OE has a Tag and return a checkbox/tag prefix."""
    tag = oe.find("one:Tag", NS)
    if tag is None:
        return ""
    index = tag.get("index", "")
    completed = tag.get("completed", "false") == "true"
    tag_def = tag_defs.get(index, {})
    tag_type = tag_def.get("type", "")
    if tag_type in ("0", "1", "2", "3", "4"):
        return "[x] " if completed else "[ ] "
    name = tag_def.get("name", "")
    if name:
        return f"[{name}] "
    return ""


def _get_list_prefix(oe, depth: int) -> str:
    """Check if an OE has a List element and return bullet/number prefix."""
    list_elem = oe.find("one:List", NS)
    if list_elem is None:
        return ""
    if list_elem.find("one:Number", NS) is not None:
        return "1. "
    if list_elem.find("one:Bullet", NS) is not None:
        return "- "
    return ""


def _process_table(table_elem) -> list[str]:
    """Convert a OneNote table to markdown table."""
    rows = table_elem.findall("one:Row", NS)
    if not rows:
        return []

    md_rows = []
    for row in rows:
        cells = row.findall("one:Cell", NS)
        cell_texts = []
        for cell in cells:
            # Collect all text in the cell
            texts = []
            for t in cell.iter():
                if _local_tag(t.tag) == "T" and t.text:
                    texts.append(_clean_text(t.text).strip())
            cell_texts.append(" ".join(texts) if texts else "")
        md_rows.append("| " + " | ".join(cell_texts) + " |")

    if len(md_rows) >= 1:
        # Insert header separator after first row
        col_count = md_rows[0].count("|") - 1
        separator = "| " + " | ".join(["---"] * col_count) + " |"
        md_rows.insert(1, separator)

    return md_rows


def _get_callback_id(img_elem) -> str | None:
    """Extract callbackID from an Image element.

    OneNote stores it as a child element: <one:CallbackID callbackID="..."/>
    not as an attribute on the Image tag itself.
    """
    # Check child element first (actual OneNote format)
    cb_elem = img_elem.find("one:CallbackID", NS)
    if cb_elem is not None:
        return cb_elem.get("callbackID")
    # Fallback: check as attribute (for compatibility)
    return img_elem.get("callbackID")


def _make_image_ref(img_elem, index: int) -> ImageRef | None:
    """Create an ImageRef from an Image element."""
    cb_id = _get_callback_id(img_elem)
    if not cb_id:
        return None

    # Try to get dimensions from Size child
    width = None
    height = None
    size = img_elem.find("one:Size", NS)
    if size is not None:
        w = size.get("width")
        h = size.get("height")
        if w:
            width = float(w)
        if h:
            height = float(h)

    return ImageRef(
        callback_id=cb_id,
        index=index,
        width=width,
        height=height,
    )


def _local_tag(tag: str) -> str:
    """Strip namespace from tag name."""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _clean_text(text: str) -> str:
    """Convert OneNote CDATA HTML to markdown-style formatting."""
    text = _convert_links(text)
    text = _convert_spans(text)
    text = re.sub(r"<br\s*/?>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = text.replace("&amp;", "&")
    text = text.replace("&lt;", "<")
    text = text.replace("&gt;", ">")
    text = text.replace("&quot;", '"')
    text = text.replace("&apos;", "'")
    text = text.replace("&nbsp;", " ")
    return text


def _convert_links(text: str) -> str:
    """Convert <a href="url">text</a> to [text](url)."""
    return re.sub(
        r'<a\s+href="([^"]*)"[^>]*>(.*?)</a>',
        r'[\2](\1)',
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )


def _convert_spans(text: str) -> str:
    """Convert styled spans to markdown formatting markers."""
    def _replace_span(m: re.Match) -> str:
        style = m.group(1)
        content = m.group(2)
        bold = "font-weight:bold" in style or "font-weight: bold" in style
        italic = "font-style:italic" in style or "font-style: italic" in style
        strike = "text-decoration:line-through" in style or "text-decoration: line-through" in style
        if bold and italic:
            return f"***{content}***"
        if bold:
            return f"**{content}**"
        if italic:
            return f"*{content}*"
        if strike:
            return f"~~{content}~~"
        return content

    return re.sub(
        r'<span\s+style="([^"]*)"[^>]*>(.*?)</span>',
        _replace_span,
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )


def parse_search_results(xml_str: str) -> list[dict]:
    """Parse FindPages result XML into a list of matches."""
    root = ET.fromstring(xml_str)
    results = []

    for nb in root.findall("one:Notebook", NS):
        nb_name = nb.get("name", "")
        for sec in nb.findall(".//one:Section", NS):
            sec_name = sec.get("name", "")
            for page in sec.findall("one:Page", NS):
                results.append({
                    "page_id": page.get("ID", ""),
                    "page_name": page.get("name", ""),
                    "notebook": nb_name,
                    "section": sec_name,
                    "last_modified": page.get("lastModifiedTime", ""),
                })

    return results
