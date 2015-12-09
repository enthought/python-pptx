# encoding: utf-8

"""
Slide and related objects.
"""

from __future__ import absolute_import

from warnings import warn

from .chart import ChartPart
from ..opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from ..oxml.parts.slide import CT_Slide
from pptx.parts.slidebase import BaseSlide
from pptx.parts.slideplaceholders import _SlidePlaceholders
from .slidenotes import SlideNotes
from ..shapes.shapetree import SlideShapeTree
from ..util import lazyproperty


class Slide(BaseSlide):
    """
    Slide part. Corresponds to package files ppt/slides/slide[1-9][0-9]*.xml.
    """
    @classmethod
    def new(cls, slide_layout, partname, package):
        """
        Return a new slide based on *slide_layout* and having *partname*,
        created from scratch.
        """
        slide_elm = CT_Slide.new()
        slide = cls(partname, CT.PML_SLIDE, slide_elm, package)
        slide.shapes.clone_layout_placeholders(slide_layout)
        slide.relate_to(slide_layout, RT.SLIDE_LAYOUT)
        return slide

    def add_chart_part(self, chart_type, chart_data):
        """
        Return the rId of a new |ChartPart| object containing a chart of
        *chart_type*, displaying *chart_data*, and related to the slide
        containing this shape tree.
        """
        chart_part = ChartPart.new(chart_type, chart_data, self.package)
        rId = self.relate_to(chart_part, RT.CHART)
        return rId

    @lazyproperty
    def placeholders(self):
        """
        Instance of |_SlidePlaceholders| containing sequence of placeholder
        shapes in this slide.
        """
        return _SlidePlaceholders(self._element.spTree, self)

    @lazyproperty
    def shapes(self):
        """
        Instance of |SlideShapeTree| containing sequence of shape objects
        appearing on this slide.
        """
        return SlideShapeTree(self)

    @property
    def slide_layout(self):
        """
        |SlideLayout| object this slide inherits appearance from.
        """
        return self.part_related_by(RT.SLIDE_LAYOUT)

    @property
    def slidelayout(self):
        """
        Deprecated. Use ``.slide_layout`` property instead.
        """
        msg = (
            'Slide.slidelayout property is deprecated. Use .slide_layout '
            'instead.'
        )
        warn(msg, UserWarning, stacklevel=2)
        return self.slide_layout

    @lazyproperty
    def notes(self):
        """Return all related notesSlides
        """
        notes_slide = None
        try:
            notes_slide = self.part_related_by(RT.NOTES_SLIDE)
        except KeyError:
            notes_slide = SlideNotes.new(self, self.package.presentation.notesMaster, self.package)
            rId = self.relate_to(notes_slide, RT.NOTES_SLIDE)
            notes_slide.relate_to(self, RT.SLIDE)

        return self.part_related_by(RT.NOTES_SLIDE)




