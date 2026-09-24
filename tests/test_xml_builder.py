"""Tests for OneNote XML builder module."""

from onenote_lib.xml_builder import (
    build_heading,
    build_list_item,
    build_outline,
    build_table,
    build_text_oe,
    cdata,
)


class TestCdata:
    def test_simple_text(self):
        assert cdata("hello") == "<![CDATA[hello]]>"

    def test_escapes_cdata_end(self):
        result = cdata("a]]>b")
        assert result == "<![CDATA[a]]]]><![CDATA[>b]]>"

    def test_empty(self):
        assert cdata("") == "<![CDATA[]]>"


class TestBuildTextOe:
    def test_plain(self):
        result = build_text_oe("hello")
        assert "<one:OE>" in result
        assert "<one:T>" in result
        assert "<![CDATA[hello]]>" in result

    def test_bold(self):
        result = build_text_oe("hello", bold=True)
        assert 'font-weight:bold' in result

    def test_italic(self):
        result = build_text_oe("hello", italic=True)
        assert 'font-style:italic' in result

    def test_bold_italic(self):
        result = build_text_oe("hello", bold=True, italic=True)
        assert 'font-weight:bold' in result
        assert 'font-style:italic' in result


class TestBuildHeading:
    def test_h1(self):
        result = build_heading("Title", 1)
        assert "font-size:20pt" in result
        assert "font-weight:bold" in result
        assert "Title" in result

    def test_h2(self):
        result = build_heading("Subtitle", 2)
        assert "font-size:16pt" in result

    def test_default_level(self):
        result = build_heading("Default")
        assert "font-size:16pt" in result


class TestBuildListItem:
    def test_bullet(self):
        result = build_list_item("item")
        assert "<one:Bullet" in result
        assert "<![CDATA[item]]>" in result

    def test_numbered(self):
        result = build_list_item("step", numbered=True)
        assert "<one:Number" in result
        assert "<![CDATA[step]]>" in result


class TestBuildTable:
    def test_simple_table(self):
        rows = [["A", "B"], ["1", "2"]]
        result = build_table(rows)
        assert "<one:Table" in result
        assert "<one:Column" in result
        assert "<one:Row>" in result
        assert "<one:Cell>" in result
        assert "<![CDATA[A]]>" in result
        assert 'hasHeaderRow="true"' in result

    def test_no_header(self):
        rows = [["A", "B"], ["1", "2"]]
        result = build_table(rows, has_header=False)
        assert 'hasHeaderRow="false"' in result

    def test_empty_rows(self):
        assert build_table([]) == ""

    def test_ragged_rows_padded(self):
        rows = [["A", "B", "C"], ["1"]]
        result = build_table(rows)
        assert result.count('index="') == 3
        assert result.count("<one:Row>") == 2


class TestBuildOutline:
    def test_wraps_elements(self):
        oes = [build_text_oe("a"), build_text_oe("b")]
        result = build_outline(oes)
        assert result.startswith("<one:Outline>")
        assert "<one:OEChildren>" in result
        assert result.endswith("</one:Outline>")
