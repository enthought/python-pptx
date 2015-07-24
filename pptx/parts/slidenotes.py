import re
from pptx.opc.constants import CONTENT_TYPE
from pptx.opc.packuri import PackURI
from pptx.shapes.shapetree import SlideShapeTree
from pptx.util import lazyproperty, Pt
from pptx.enum.shapes import PP_PLACEHOLDER_TYPE
from pptx.parts.slideplaceholders import _SlidePlaceholders
from pptx.parts.slidebase import BaseSlide
from pptx.oxml.parts.slidenotes import CT_SlideNotes


class SlideNotes(BaseSlide):
  """This class will represent the Part of the notesSlide. Any notes retrieved
  from the presentation slides will be an instance of this class.
  """

  @classmethod
  def new(cls, slide, partname, package):
    notes_slide_elm = CT_SlideNotes.new()
    partname = PackURI(re.sub("slide", "notesSlide", partname))
    slide = cls(partname, CONTENT_TYPE.PML_NOTES_SLIDE, notes_slide_elm, package)
    return slide

  @lazyproperty
  def shapes(self):
    """
    Instance of |_SlideShapeTree| containing sequence of shape objects
    appearing on this slide.
    """
    return SlideShapeTree(self)

  @lazyproperty
  def placeholders(self):
    """
    Instance of |_SlidePlaceholders| containing sequence of placeholder
    shapes in this slide.
    """
    return _SlidePlaceholders(self)

  def add_multiline_note(self, text):
    for line in text.split('\n'):
      self.add_note(line)

  def add_note(self, text):
    """Add some text to the notesSlide, return paragraph
    that was added or False if no textframes were found
    """
    for shape in self.shapes:
      if shape.has_text_frame and shape.is_placeholder:
        para = shape.text_frame.add_paragraph()
        para.text = text
    return False

  def clear_notes(self):
    """Remove all current notes from the slide
    """
    for shape in self.shapes:
      if shape.has_text_frame:
        shape.text_frame.clear()

  def get_slide_runs(self):
    for shape in self.shapes:
      if shape.has_text_frame and shape.is_placeholder:
        for p in shape.text_frame.paragraphs:
          for run in p.runs:
            yield run
