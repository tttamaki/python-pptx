"""Custom element classes for Markup Compatibility (mc:) elements.

Handles mc:AlternateContent, mc:Choice, and mc:Fallback as defined in
ECMA-376 Part 3 (Markup Compatibility and Extensibility).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import BaseOxmlElement

if TYPE_CHECKING:
    from pptx.oxml.shapes import ShapeElement


class CT_Choice(BaseOxmlElement):
    """`mc:Choice` element — the preferred content branch of AlternateContent."""


class CT_Fallback(BaseOxmlElement):
    """`mc:Fallback` element — the fallback content branch of AlternateContent."""


class CT_AlternateContent(BaseOxmlElement):
    """`mc:AlternateContent` wrapper that exposes the first `mc:Choice` child's shapes.

    When iterating shape elements on a slide, AlternateContent is treated
    transparently: the shapes inside the first mc:Choice are yielded as if
    they were direct children of the shape-tree.
    """

    _shape_tags = (
        qn("p:sp"),
        qn("p:grpSp"),
        qn("p:graphicFrame"),
        qn("p:cxnSp"),
        qn("p:pic"),
        qn("p:contentPart"),
    )

    def iter_shape_elms(self) -> Iterator[ShapeElement]:
        """Yield shape child elements from the first `mc:Choice` child.

        Falls back to `mc:Fallback` shapes if no Choice is present.
        """
        choice = self.find(qn("mc:Choice"))
        source = choice if choice is not None else self.find(qn("mc:Fallback"))
        if source is None:
            return
        for elm in source:
            if elm.tag in self._shape_tags:
                yield elm
