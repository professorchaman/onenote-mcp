"""Unit tests for OneNote XML parser."""

import pytest

from onenote_lib.xml_parser import (
    parse_notebooks,
    parse_page_to_markdown,
    parse_search_results,
)

NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


HIERARCHY_XML = f"""<?xml version="1.0"?>
<one:Notebooks xmlns:one="{NS}">
  <one:Notebook name="Work Notes" ID="nb-001" path="C:\\Users\\test\\Work Notes"
                lastModifiedTime="2026-02-15T10:00:00Z">
    <one:Section name="Meeting Notes" ID="sec-001" path="C:\\Users\\test\\Work Notes\\Meeting Notes.one">
      <one:Page ID="page-001" name="Monday Standup" lastModifiedTime="2026-02-14T09:00:00Z" pageLevel="0"/>
      <one:Page ID="page-002" name="Sprint Review" lastModifiedTime="2026-02-13T14:00:00Z" pageLevel="0"/>
    </one:Section>
    <one:SectionGroup name="Archive" ID="sg-001">
      <one:Section name="Old Notes" ID="sec-002">
        <one:Page ID="page-003" name="Archived Page" pageLevel="0"/>
      </one:Section>
    </one:SectionGroup>
    <one:SectionGroup name="Recycle Bin" ID="sg-bin" isRecycleBin="true">
      <one:Section name="Deleted" ID="sec-deleted"/>
    </one:SectionGroup>
  </one:Notebook>
  <one:Notebook name="Personal" ID="nb-002" path="C:\\Users\\test\\Personal"
                lastModifiedTime="2026-02-10T08:00:00Z">
    <one:Section name="Journal" ID="sec-003">
      <one:Page ID="page-004" name="Feb 10" pageLevel="0"/>
    </one:Section>
  </one:Notebook>
</one:Notebooks>"""


PAGE_XML = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="page-001" name="Test Page">
  <one:Title>
    <one:OE><one:T><![CDATA[Test Page]]></one:T></one:OE>
  </one:Title>
  <one:Outline>
    <one:OEChildren>
      <one:OE>
        <one:T><![CDATA[This is paragraph one.]]></one:T>
      </one:OE>
      <one:OE>
        <one:T><![CDATA[This is paragraph two with <b>bold</b> text.]]></one:T>
      </one:OE>
      <one:OE>
        <one:Image>
          <one:Size width="640" height="480" isSetByUser="true"/>
          <one:CallbackID callbackID="img-001"/>
        </one:Image>
      </one:OE>
      <one:OE>
        <one:T><![CDATA[Text after image.]]></one:T>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
  <one:Outline>
    <one:OEChildren>
      <one:OE>
        <one:Table>
          <one:Row>
            <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Header 1]]></one:T></one:OE></one:OEChildren></one:Cell>
            <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Header 2]]></one:T></one:OE></one:OEChildren></one:Cell>
          </one:Row>
          <one:Row>
            <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Data 1]]></one:T></one:OE></one:OEChildren></one:Cell>
            <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Data 2]]></one:T></one:OE></one:OEChildren></one:Cell>
          </one:Row>
        </one:Table>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>"""


PAGE_NO_IMAGES_XML = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="page-005" name="Plain Page">
  <one:Title>
    <one:OE><one:T><![CDATA[Plain Page]]></one:T></one:OE>
  </one:Title>
  <one:Outline>
    <one:OEChildren>
      <one:OE>
        <one:T><![CDATA[Just text, no images.]]></one:T>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>"""


SEARCH_XML = f"""<?xml version="1.0"?>
<one:Notebooks xmlns:one="{NS}">
  <one:Notebook name="Work Notes" ID="nb-001">
    <one:Section name="Meeting Notes" ID="sec-001">
      <one:Page ID="page-001" name="Monday Standup" lastModifiedTime="2026-02-14T09:00:00Z"/>
    </one:Section>
  </one:Notebook>
</one:Notebooks>"""


class TestParseNotebooks:
    def test_basic_parsing(self):
        notebooks = parse_notebooks(HIERARCHY_XML)
        assert len(notebooks) == 2
        assert notebooks[0].name == "Work Notes"
        assert notebooks[0].id == "nb-001"
        assert notebooks[1].name == "Personal"

    def test_sections(self):
        notebooks = parse_notebooks(HIERARCHY_XML)
        nb = notebooks[0]
        assert len(nb.sections) == 1
        assert nb.sections[0].name == "Meeting Notes"
        assert len(nb.sections[0].pages) == 2

    def test_section_groups(self):
        notebooks = parse_notebooks(HIERARCHY_XML)
        nb = notebooks[0]
        # Should have Archive but NOT Recycle Bin
        assert len(nb.section_groups) == 1
        assert nb.section_groups[0].name == "Archive"
        assert len(nb.section_groups[0].sections) == 1

    def test_pages(self):
        notebooks = parse_notebooks(HIERARCHY_XML)
        pages = notebooks[0].sections[0].pages
        assert len(pages) == 2
        assert pages[0].name == "Monday Standup"
        assert pages[0].id == "page-001"

    def test_nested_section_group_pages(self):
        notebooks = parse_notebooks(HIERARCHY_XML)
        sg = notebooks[0].section_groups[0]
        assert sg.sections[0].pages[0].name == "Archived Page"


class TestParsePageToMarkdown:
    def test_basic_content(self):
        md, images = parse_page_to_markdown(PAGE_XML)
        assert "# Test Page" in md
        assert "This is paragraph one." in md
        assert "This is paragraph two with bold text." in md
        assert "Text after image." in md

    def test_image_references(self):
        md, images = parse_page_to_markdown(PAGE_XML)
        assert len(images) == 1
        assert images[0].callback_id == "img-001"
        assert images[0].width == 640.0
        assert images[0].height == 480.0
        assert "[Image 1]" in md

    def test_table_parsing(self):
        md, _ = parse_page_to_markdown(PAGE_XML)
        assert "Header 1" in md
        assert "Header 2" in md
        assert "Data 1" in md
        assert "|" in md
        assert "---" in md

    def test_no_images(self):
        md, images = parse_page_to_markdown(PAGE_NO_IMAGES_XML)
        assert len(images) == 0
        assert "Just text, no images." in md

    def test_html_stripping(self):
        md, _ = parse_page_to_markdown(PAGE_XML)
        # <b> tags should be stripped
        assert "<b>" not in md
        assert "bold" in md


class TestParseSearchResults:
    def test_basic_search(self):
        results = parse_search_results(SEARCH_XML)
        assert len(results) == 1
        assert results[0]["page_name"] == "Monday Standup"
        assert results[0]["notebook"] == "Work Notes"
        assert results[0]["section"] == "Meeting Notes"

    def test_empty_search(self):
        empty_xml = f'<one:Notebooks xmlns:one="{NS}"/>'
        results = parse_search_results(empty_xml)
        assert len(results) == 0


class TestFormattingPreservation:
    """Phase 4: formatting spans should convert to markdown markers."""

    def test_bold_preserved(self):
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="p-fmt" name="Fmt">
  <one:Title><one:OE><one:T><![CDATA[Fmt]]></one:T></one:OE></one:Title>
  <one:Outline><one:OEChildren>
    <one:OE><one:T><![CDATA[<span style="font-weight:bold">important</span>]]></one:T></one:OE>
  </one:OEChildren></one:Outline>
</one:Page>"""
        md, _ = parse_page_to_markdown(xml)
        assert "**important**" in md

    def test_italic_preserved(self):
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="p-it" name="It">
  <one:Title><one:OE><one:T><![CDATA[It]]></one:T></one:OE></one:Title>
  <one:Outline><one:OEChildren>
    <one:OE><one:T><![CDATA[<span style="font-style:italic">emphasis</span>]]></one:T></one:OE>
  </one:OEChildren></one:Outline>
</one:Page>"""
        md, _ = parse_page_to_markdown(xml)
        assert "*emphasis*" in md

    def test_strikethrough_preserved(self):
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="p-st" name="St">
  <one:Title><one:OE><one:T><![CDATA[St]]></one:T></one:OE></one:Title>
  <one:Outline><one:OEChildren>
    <one:OE><one:T><![CDATA[<span style="text-decoration:line-through">removed</span>]]></one:T></one:OE>
  </one:OEChildren></one:Outline>
</one:Page>"""
        md, _ = parse_page_to_markdown(xml)
        assert "~~removed~~" in md

    def test_hyperlink_preserved(self):
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="p-lk" name="Lk">
  <one:Title><one:OE><one:T><![CDATA[Lk]]></one:T></one:OE></one:Title>
  <one:Outline><one:OEChildren>
    <one:OE><one:T><![CDATA[Visit <a href="https://example.com">here</a> now]]></one:T></one:OE>
  </one:OEChildren></one:Outline>
</one:Page>"""
        md, _ = parse_page_to_markdown(xml)
        assert "[here](https://example.com)" in md


class TestTagsAndCheckboxes:
    """Phase 4: tag/checkbox rendering."""

    def test_todo_unchecked(self):
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="p-td" name="Td">
  <one:TagDef index="0" name="To Do" type="0" />
  <one:Title><one:OE><one:T><![CDATA[Td]]></one:T></one:OE></one:Title>
  <one:Outline><one:OEChildren>
    <one:OE>
      <one:Tag index="0" completed="false" />
      <one:T><![CDATA[Buy milk]]></one:T>
    </one:OE>
  </one:OEChildren></one:Outline>
</one:Page>"""
        md, _ = parse_page_to_markdown(xml)
        assert "[ ] Buy milk" in md

    def test_todo_checked(self):
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="p-tc" name="Tc">
  <one:TagDef index="0" name="To Do" type="0" />
  <one:Title><one:OE><one:T><![CDATA[Tc]]></one:T></one:OE></one:Title>
  <one:Outline><one:OEChildren>
    <one:OE>
      <one:Tag index="0" completed="true" />
      <one:T><![CDATA[Done task]]></one:T>
    </one:OE>
  </one:OEChildren></one:Outline>
</one:Page>"""
        md, _ = parse_page_to_markdown(xml)
        assert "[x] Done task" in md


class TestListHandling:
    """Phase 4: bullet and numbered list rendering."""

    def test_bullet_list(self):
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="p-bl" name="Bl">
  <one:Title><one:OE><one:T><![CDATA[Bl]]></one:T></one:OE></one:Title>
  <one:Outline><one:OEChildren>
    <one:OE>
      <one:List><one:Bullet bullet="2" fontSize="11.0" /></one:List>
      <one:T><![CDATA[Item one]]></one:T>
    </one:OE>
    <one:OE>
      <one:List><one:Bullet bullet="2" fontSize="11.0" /></one:List>
      <one:T><![CDATA[Item two]]></one:T>
    </one:OE>
  </one:OEChildren></one:Outline>
</one:Page>"""
        md, _ = parse_page_to_markdown(xml)
        assert "- Item one" in md
        assert "- Item two" in md

    def test_numbered_list(self):
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="p-nl" name="Nl">
  <one:Title><one:OE><one:T><![CDATA[Nl]]></one:T></one:OE></one:Title>
  <one:Outline><one:OEChildren>
    <one:OE>
      <one:List><one:Number numberSequence="1" numberFormat="##." /></one:List>
      <one:T><![CDATA[First]]></one:T>
    </one:OE>
  </one:OEChildren></one:Outline>
</one:Page>"""
        md, _ = parse_page_to_markdown(xml)
        assert "1. First" in md


class TestNoDuplicateText:
    """Regression tests: table cell text must not appear as standalone lines."""

    def test_table_text_not_duplicated(self):
        md, _ = parse_page_to_markdown(PAGE_XML)
        lines = md.split("\n")
        header_standalone = [l for l in lines if l.strip() == "Header 1"]
        data_standalone = [l for l in lines if l.strip() == "Data 1"]
        assert len(header_standalone) == 0, "Table header text duplicated as standalone line"
        assert len(data_standalone) == 0, "Table data text duplicated as standalone line"
        assert "| Header 1 | Header 2 |" in md
        assert "| Data 1 | Data 2 |" in md

    def test_mixed_content_no_duplication(self):
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="p-mix" name="Mixed">
  <one:Title><one:OE><one:T><![CDATA[Mixed]]></one:T></one:OE></one:Title>
  <one:Outline><one:OEChildren>
    <one:OE><one:T><![CDATA[Before table.]]></one:T></one:OE>
    <one:OE>
      <one:Table>
        <one:Row>
          <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Cell A]]></one:T></one:OE></one:OEChildren></one:Cell>
        </one:Row>
      </one:Table>
    </one:OE>
    <one:OE><one:T><![CDATA[After table.]]></one:T></one:OE>
  </one:OEChildren></one:Outline>
</one:Page>"""
        md, _ = parse_page_to_markdown(xml)
        assert md.count("Cell A") == 1, f"Cell text appeared {md.count('Cell A')} times"
        assert "Before table." in md
        assert "After table." in md
