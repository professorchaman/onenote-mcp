"""Tests for markdown-to-OneNote converter."""

from onenote_lib.markdown_to_onenote import markdown_to_outline


class TestMarkdownToOutline:
    def test_plain_paragraph(self):
        result = markdown_to_outline("Hello world")
        assert "<one:Outline>" in result
        assert "<![CDATA[Hello world]]>" in result

    def test_heading(self):
        result = markdown_to_outline("## My Heading")
        assert "font-size:16pt" in result
        assert "My Heading" in result

    def test_h1(self):
        result = markdown_to_outline("# Top Level")
        assert "font-size:20pt" in result

    def test_bullet_list(self):
        md = "- first\n- second"
        result = markdown_to_outline(md)
        assert result.count("<one:Bullet") == 2
        assert "<![CDATA[first]]>" in result
        assert "<![CDATA[second]]>" in result

    def test_numbered_list(self):
        md = "1. alpha\n2. beta"
        result = markdown_to_outline(md)
        assert result.count("<one:Number") == 2

    def test_bold_line(self):
        result = markdown_to_outline("**important**")
        assert "font-weight:bold" in result
        assert "important" in result

    def test_italic_line(self):
        result = markdown_to_outline("*emphasis*")
        assert "font-style:italic" in result

    def test_table(self):
        md = "| A | B |\n| --- | --- |\n| 1 | 2 |"
        result = markdown_to_outline(md)
        assert "<one:Table" in result
        assert "<![CDATA[A]]>" in result
        assert "<![CDATA[1]]>" in result
        assert 'hasHeaderRow="true"' in result

    def test_mixed_content(self):
        md = "# Title\n\nSome text\n\n- bullet\n\n| H |\n| --- |\n| D |"
        result = markdown_to_outline(md)
        assert "font-size:20pt" in result
        assert "<![CDATA[Some text]]>" in result
        assert "<one:Bullet" in result
        assert "<one:Table" in result

    def test_empty_lines_skipped(self):
        md = "\n\ntext\n\n"
        result = markdown_to_outline(md)
        assert "<![CDATA[text]]>" in result

    def test_empty_string(self):
        result = markdown_to_outline("")
        assert "<one:Outline>" in result
        assert "<one:OEChildren>" in result
