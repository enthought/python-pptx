"""
Example of accessing the notes slides of a presentation.
Requires python-pptx 0.5.6 or later.

ryan@ryanday.net
"""
from pptx.oxml.xmlchemy import BaseOxmlElement, OneAndOnlyOne, ZeroOrOne
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls


"""
http://msdn.microsoft.com/en-us/library/office/gg278319%28v=office.15%29.aspx
"""


class CT_SlideNotes(BaseOxmlElement):
    # This is a hack.
    notes_id = 1000
    """
    ``<p:notes>`` element, root of a notesSlide part
    """
    cSld = OneAndOnlyOne('p:cSld')
    clrMapOvr = ZeroOrOne('p:clrMapOvr', successors=(
        'p:transition', 'p:timing', 'p:extLst'
    ))

    @classmethod
    def new(cls):
        """
        Return a new ``<p:notes>`` element configured as a base slide shape.
        """
        return parse_xml(cls._notes_xml())

    @staticmethod
    def _notes_xml():
      """From http://msdn.microsoft.com/en-us/library/office/gg278319%28v=office.15%29.aspx#sectionSection4
      """
      return (
        '<p:notes %s>\n'
        '  <p:cSld>\n'
        '    <p:spTree>\n'
        '      <p:nvGrpSpPr>\n'
        '        <p:cNvPr id="1" name=""/>\n'
        '        <p:cNvGrpSpPr/>\n'
        '        <p:nvPr/>\n'
        '      </p:nvGrpSpPr>\n'
        '      <p:grpSpPr/>\n'
        '    </p:spTree>\n'
        '  </p:cSld>\n'
        '  <p:clrMapOvr>\n'
        '    <a:masterClrMapping/>\n'
        '  </p:clrMapOvr>\n'
        '</p:notes>\n' % nsdecls('p', 'a', 'r')
      )
