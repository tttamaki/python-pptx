"""Custom element classes for Office Math (m:) and drawing math (a14:m) elements.

Handles the a14:m wrapper (Office 2010 inline math) and the core m:oMath /
m:r / m:t elements from OOXML math namespace.
"""

from __future__ import annotations

from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import BaseOxmlElement


class CT_OMathText(BaseOxmlElement):
    """`m:t` element — holds the literal text of a math run."""


class CT_OMathRun(BaseOxmlElement):
    """`m:r` element — a single math run, containing an optional `m:t` child."""

    @property
    def text(self) -> str:  # pyright: ignore[reportIncompatibleMethodOverride]
        """Text from the `m:t` child element, empty string when absent."""
        t = self.find(qn("m:t"))
        if t is None:
            return ""
        return t.text or ""  # type: ignore[return-value]


class CT_OMath(BaseOxmlElement):
    """`m:oMath` element — an inline math expression.

    Yields concatenated text of all `m:r/m:t` descendants.
    """

    @property
    def text(self) -> str:  # pyright: ignore[reportIncompatibleMethodOverride]
        """Concatenated text of all `m:r` child elements."""
        parts: list[str] = []
        for r in self.iterchildren(qn("m:r")):
            t = r.find(qn("m:t"))
            if t is not None and t.text:
                parts.append(t.text)
        return "".join(parts)


class CT_OfficeMath(BaseOxmlElement):
    """`a14:m` element — Office 2010 drawing math wrapper.

    Acts as an inline run-like element inside `a:p`, wrapping `m:oMath`.
    Exposes a `.text` property for compatibility with the paragraph text
    extraction logic, returning the concatenated math text.
    """

    @property
    def text(self) -> str:  # pyright: ignore[reportIncompatibleMethodOverride]
        """Math text from the contained `m:oMath` child, empty string when absent."""
        oMath = self.find(qn("m:oMath"))
        if oMath is None:
            return ""
        # oMath may be a generic BaseOxmlElement if not yet registered; handle both
        if hasattr(oMath, "text") and callable(getattr(type(oMath), "text", None)):
            # registered CT_OMath instance
            return oMath.text  # type: ignore[return-value]
        # fallback: concatenate m:r/m:t text manually
        parts: list[str] = []
        for r in oMath.iterchildren(qn("m:r")):
            t = r.find(qn("m:t"))
            if t is not None and t.text:
                parts.append(t.text)
        return "".join(parts)
