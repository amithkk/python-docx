"""Paragraph-related proxy types."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, List, cast

import datetime as dt

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.shared import StoryChild
from docx.styles.style import ParagraphStyle
from docx.text.hyperlink import Hyperlink
from docx.text.pagebreak import RenderedPageBreak
from docx.text.parfmt import ParagraphFormat
from docx.text.run import Run

if TYPE_CHECKING:
    import docx.types as t
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml.text.paragraph import CT_P
    from docx.revision import TrackedDeletion, TrackedInsertion
    from docx.styles.style import CharacterStyle


class Paragraph(StoryChild):
    """Proxy object wrapping a `<w:p>` element."""

    def __init__(self, p: CT_P, parent: t.ProvidesStoryPart):
        super(Paragraph, self).__init__(parent)
        self._p = self._element = p

    def add_run(self, text: str | None = None, style: str | CharacterStyle | None = None) -> Run:
        """Append run containing `text` and having character-style `style`.

        `text` can contain tab (``\\t``) characters, which are converted to the
        appropriate XML form for a tab. `text` can also include newline (``\\n``) or
        carriage return (``\\r``) characters, each of which is converted to a line
        break. When `text` is `None`, the new run is empty.
        """
        r = self._p.add_r()
        run = Run(r, self)
        if text:
            run.text = text
        if style:
            run.style = style
        return run

    def add_run_tracked(
        self,
        text: str | None = None,
        style: str | CharacterStyle | None = None,
        author: str = "",
        revision_id: int | None = None,
    ) -> TrackedInsertion:
        """Append a tracked insertion containing a run with the specified text.

        The run is wrapped in a `w:ins` element, marking it as inserted content
        when track changes is enabled.

        Args:
            text: Text to add to the run.
            style: Character style to apply to the run.
            author: Author name for the revision. Defaults to empty string.
            revision_id: Unique ID for this revision. Auto-generated if not provided.

        Returns:
            A TrackedInsertion object wrapping the `w:ins` element.
        """
        from docx.oxml.parser import OxmlElement
        from docx.revision import TrackedInsertion

        if revision_id is None:
            revision_id = self._next_revision_id()

        ins = OxmlElement(
            "w:ins",
            attrs={
                qn("w:id"): str(revision_id),
                qn("w:author"): author,
                qn("w:date"): dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            },
        )

        r = OxmlElement("w:r")
        ins.append(r)
        self._p.append(ins)  # pyright: ignore[reportUnknownMemberType]

        tracked_insertion = TrackedInsertion(ins, self)  # pyright: ignore[reportArgumentType]
        if text:
            for run in tracked_insertion.runs:
                run.text = text
        if style:
            for run in tracked_insertion.runs:
                run.style = style

        return tracked_insertion

    def _next_revision_id(self) -> int:
        """Generate the next unique revision ID for this document."""
        max_id = 0
        for ins_or_del in self._p.xpath("//w:ins | //w:del"):
            id_val = ins_or_del.get(qn("w:id"))  # pyright: ignore[reportUnknownMemberType]
            if id_val is not None:
                try:
                    max_id = max(max_id, int(id_val))
                except ValueError:
                    pass
        return max_id + 1

    @property
    def alignment(self) -> WD_PARAGRAPH_ALIGNMENT | None:
        """A member of the :ref:`WdParagraphAlignment` enumeration specifying the
        justification setting for this paragraph.

        A value of |None| indicates the paragraph has no directly-applied alignment
        value and will inherit its alignment value from its style hierarchy. Assigning
        |None| to this property removes any directly-applied alignment value.
        """
        return self._p.alignment

    @alignment.setter
    def alignment(self, value: WD_PARAGRAPH_ALIGNMENT):
        self._p.alignment = value

    def clear(self):
        """Return this same paragraph after removing all its content.

        Paragraph-level formatting, such as style, is preserved.
        """
        self._p.clear_content()
        return self

    @property
    def contains_page_break(self) -> bool:
        """`True` when one or more rendered page-breaks occur in this paragraph."""
        return bool(self._p.lastRenderedPageBreaks)

    @property
    def hyperlinks(self) -> List[Hyperlink]:
        """A |Hyperlink| instance for each hyperlink in this paragraph."""
        return [Hyperlink(hyperlink, self) for hyperlink in self._p.hyperlink_lst]

    def insert_paragraph_before(
        self, text: str | None = None, style: str | ParagraphStyle | None = None
    ) -> Paragraph:
        """Return a newly created paragraph, inserted directly before this paragraph.

        If `text` is supplied, the new paragraph contains that text in a single run. If
        `style` is provided, that style is assigned to the new paragraph.
        """
        paragraph = self._insert_paragraph_before()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    def iter_inner_content(
        self, include_revisions: bool = False
    ) -> Iterator[Run | Hyperlink | TrackedInsertion | TrackedDeletion]:
        """Generate the runs and hyperlinks in this paragraph, in the order they appear.

        The content in a paragraph consists of both runs and hyperlinks. This method
        allows accessing each of those separately, in document order, for when the
        precise position of the hyperlink within the paragraph text is important. Note
        that a hyperlink itself contains runs.

        Args:
            include_revisions: If True, also yields `TrackedInsertion` and
                `TrackedDeletion` objects for run-level tracked changes
                (`w:ins` and `w:del` elements that wrap runs).
                Defaults to False for backward compatibility.

        Yields:
            Run, Hyperlink, TrackedInsertion, or TrackedDeletion objects in
            document order.
        """
        from docx.revision import TrackedDeletion, TrackedInsertion

        if include_revisions:
            elements = self._p.inner_content_with_revisions
        else:
            elements = self._p.inner_content_elements

        for element in elements:
            tag = element.tag  # pyright: ignore[reportUnknownMemberType]
            if tag == qn("w:r"):
                yield Run(element, self)
            elif tag == qn("w:hyperlink"):
                yield Hyperlink(element, self)  # pyright: ignore[reportArgumentType]
            elif tag == qn("w:ins"):
                yield TrackedInsertion(element, self)  # pyright: ignore[reportArgumentType]
            elif tag == qn("w:del"):
                yield TrackedDeletion(element, self)  # pyright: ignore[reportArgumentType]

    @property
    def paragraph_format(self):
        """The |ParagraphFormat| object providing access to the formatting properties
        for this paragraph, such as line spacing and indentation."""
        return ParagraphFormat(self._element)

    @property
    def rendered_page_breaks(self) -> List[RenderedPageBreak]:
        """All rendered page-breaks in this paragraph.

        Most often an empty list, sometimes contains one page-break, but can contain
        more than one is rare or contrived cases.
        """
        return [RenderedPageBreak(lrpb, self) for lrpb in self._p.lastRenderedPageBreaks]

    @property
    def runs(self) -> List[Run]:
        """Sequence of |Run| instances corresponding to the <w:r> elements in this
        paragraph."""
        return [Run(r, self) for r in self._p.r_lst]

    @property
    def style(self) -> ParagraphStyle | None:
        """Read/Write.

        |_ParagraphStyle| object representing the style assigned to this paragraph. If
        no explicit style is assigned to this paragraph, its value is the default
        paragraph style for the document. A paragraph style name can be assigned in lieu
        of a paragraph style object. Assigning |None| removes any applied style, making
        its effective value the default paragraph style for the document.
        """
        style_id = self._p.style
        style = self.part.get_style(style_id, WD_STYLE_TYPE.PARAGRAPH)
        return cast(ParagraphStyle, style)

    @style.setter
    def style(self, style_or_name: str | ParagraphStyle | None):
        style_id = self.part.get_style_id(style_or_name, WD_STYLE_TYPE.PARAGRAPH)
        self._p.style = style_id

    @property
    def text(self) -> str:
        """The textual content of this paragraph.

        The text includes the visible-text portion of any hyperlinks in the paragraph.
        Tabs and line breaks in the XML are mapped to ``\\t`` and ``\\n`` characters
        respectively.

        Assigning text to this property causes all existing paragraph content to be
        replaced with a single run containing the assigned text. A ``\\t`` character in
        the text is mapped to a ``<w:tab/>`` element and each ``\\n`` or ``\\r``
        character is mapped to a line break. Paragraph-level formatting, such as style,
        is preserved. All run-level formatting, such as bold or italic, is removed.
        """
        return self._p.text

    @text.setter
    def text(self, text: str | None):
        self.clear()
        self.add_run(text)

    def replace_tracked(
        self,
        search_text: str,
        replace_text: str,
        author: str = "",
        comment: str | None = None,
    ) -> int:
        """Replace all occurrences of `search_text` with `replace_text` using track changes.

        Each replacement creates a tracked deletion of `search_text` and a tracked
        insertion of `replace_text`. If `comment` is provided, a comment is attached
        to the replacement text explaining the change.

        Args:
            search_text: Text to find and replace.
            replace_text: Text to insert in place of search_text.
            author: Author name for the revision. Defaults to empty string.
            comment: Optional comment text to attach to each replacement.

        Returns:
            The number of replacements made.
        """
        from docx.oxml.parser import OxmlElement
        from docx.oxml.text.run import CT_R
        from typing import cast as typing_cast

        count = 0
        now = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        runs = list(self.runs)
        for run in runs:
            text = run.text
            if search_text not in text:
                continue

            parts = text.split(search_text)

            r_elem = run._r
            parent = r_elem.getparent()
            if parent is None:
                continue
            index = list(parent).index(r_elem)

            parent.remove(r_elem)

            insert_idx = index
            for i, part in enumerate(parts):
                if part:
                    new_r = OxmlElement("w:r")
                    new_t = OxmlElement("w:t")
                    new_t.text = part
                    if part.startswith(" ") or part.endswith(" "):
                        new_t.set(qn("xml:space"), "preserve")
                    new_r.append(new_t)
                    parent.insert(insert_idx, new_r)
                    insert_idx += 1

                if i < len(parts) - 1:
                    rev_id = self._next_revision_id()

                    del_elem = OxmlElement(
                        "w:del",
                        attrs={
                            qn("w:id"): str(rev_id),
                            qn("w:author"): author,
                            qn("w:date"): now,
                        },
                    )
                    del_r = OxmlElement("w:r")
                    del_text = OxmlElement("w:delText")
                    del_text.text = search_text
                    del_r.append(del_text)
                    del_elem.append(del_r)
                    parent.insert(insert_idx, del_elem)
                    insert_idx += 1

                    comment_id = None
                    if comment:
                        doc_comments = self.part._document_part.comments  # pyright: ignore[reportAttributeAccessIssue]
                        comment_obj = doc_comments.add_comment(text=comment, author=author)
                        comment_id = comment_obj.comment_id

                        comment_start = OxmlElement(
                            "w:commentRangeStart", attrs={qn("w:id"): str(comment_id)}
                        )
                        parent.insert(insert_idx, comment_start)
                        insert_idx += 1

                    rev_id = self._next_revision_id()
                    ins_elem = OxmlElement(
                        "w:ins",
                        attrs={
                            qn("w:id"): str(rev_id),
                            qn("w:author"): author,
                            qn("w:date"): now,
                        },
                    )
                    ins_r = OxmlElement("w:r")
                    ins_t = OxmlElement("w:t")
                    ins_t.text = replace_text
                    ins_r.append(ins_t)
                    ins_elem.append(ins_r)
                    parent.insert(insert_idx, ins_elem)
                    insert_idx += 1

                    if comment_id is not None:
                        comment_end = OxmlElement(
                            "w:commentRangeEnd", attrs={qn("w:id"): str(comment_id)}
                        )
                        parent.insert(insert_idx, comment_end)
                        insert_idx += 1

                        comment_ref_run = typing_cast(CT_R, OxmlElement("w:r"))
                        comment_ref_rPr = comment_ref_run.get_or_add_rPr()
                        comment_ref_rPr.style = "CommentReference"
                        comment_ref_run.append(
                            OxmlElement("w:commentReference", attrs={qn("w:id"): str(comment_id)})
                        )
                        parent.insert(insert_idx, comment_ref_run)
                        insert_idx += 1

                    count += 1

        return count

    def _insert_paragraph_before(self):
        """Return a newly created paragraph, inserted directly before this paragraph."""
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)
