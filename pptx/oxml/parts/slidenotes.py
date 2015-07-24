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
        return (
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
                   '    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
                   '    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
                   '    xmlns:mv="urn:schemas-microsoft-com:mac:vml"'
                   '    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
                   '    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
                   '    xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
                   '    xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml"'
                   '    xmlns:pvml="urn:schemas-microsoft-com:office:powerpoint"'
                   '    xmlns:com="http://schemas.openxmlformats.org/drawingml/2006/compatibility"'
                   '    xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" showMasterPhAnim="0" showMasterSp="0">'
                   '<p:cSld>'
                   '<p:spTree>'
                   '<p:nvGrpSpPr>'
                   '<p:cNvPr id="52" name="Shape 52"/>'
                   '<p:cNvGrpSpPr/>'
                   '<p:nvPr/>'
                   '</p:nvGrpSpPr>'
                   '<p:grpSpPr>'
                   '<a:xfrm>'
                   '<a:off x="0" y="0"/>'
                   '<a:ext cx="0" cy="0"/>'
                   '<a:chOff x="0" y="0"/>'
                   '<a:chExt cx="0" cy="0"/>'
                   '</a:xfrm>'
                   '</p:grpSpPr>'
                   '<p:sp>'
                   '  <p:nvSpPr>'
                   '    <p:cNvPr id="53" name="Shape 53"/>'
                   '    <p:cNvSpPr/>'
                   # '    <p:nvPr>'
                   # '      <p:ph idx="2" type="sldImg"/>'
                   # '    </p:nvPr>'
                   '  </p:nvSpPr>'
                   '  <p:spPr>'
                   '<a:xfrm>'
                   '<a:off x="381300" y="685800"/>'
                   '<a:ext cx="6096075" cy="3429000"/>'
                   '</a:xfrm>'
                   '<a:custGeom>'
                   '<a:pathLst>'
                   '<a:path extrusionOk="0" h="120000" w="120000">'
                   '<a:moveTo>'
                   '<a:pt x="0" y="0"/>'
                   '</a:moveTo>'
                   '<a:lnTo>'
                   '<a:pt x="120000" y="0"/>'
                   '</a:lnTo>'
                   '<a:lnTo>'
                   '<a:pt x="120000" y="120000"/>'
                   '</a:lnTo>'
                   '<a:lnTo>'
                   '<a:pt x="0" y="120000"/>'
                   '</a:lnTo>'
                   '<a:close/>'
                   '</a:path>'
                   '</a:pathLst>'
                   '</a:custGeom>'
                   '<a:noFill/>'
                   '<a:ln cap="flat" cmpd="sng" w="9525">'
                   '<a:solidFill>'
                   '<a:srgbClr val="000000"/>'
                   '</a:solidFill>'
                   '<a:prstDash val="solid"/>'
                   '<a:round/>'
                   '<a:headEnd len="med" w="med" type="none"/>'
                   '<a:tailEnd len="med" w="med" type="none"/>'
                   '</a:ln>'
                   '</p:spPr>'
                   '</p:sp>'
                   '<p:sp>'
                   '<p:nvSpPr>'
                   '<p:cNvPr id="54" name="Shape 54"/>'
                   '<p:cNvSpPr txBox="1"/>'
                   '  <p:nvPr>'
                   '    <p:ph idx="1" type="body"/>'
                   '  </p:nvPr>'
                   '</p:nvSpPr>'
                   '<p:spPr>'
                   '<a:xfrm>'
                   '<a:off x="685800" y="4343400"/>'
                   '<a:ext cx="5486399" cy="4114800"/>'
                   '</a:xfrm>'
                   '<a:prstGeom prst="rect">'
                   '<a:avLst/>'
                   '</a:prstGeom>'
                   '</p:spPr>'
                   '<p:txBody>'
                   '<a:bodyPr anchorCtr="0" anchor="t" bIns="91425" lIns="91425" rIns="91425" tIns="91425">'
                   '<a:noAutofit/>'
                   '</a:bodyPr>'
                   '<a:lstStyle/>'
                   '</p:txBody>'
                   '</p:sp>'
                   '</p:spTree>'
                   '</p:cSld>'
                   '<p:clrMapOvr>'
                   '<a:masterClrMapping/>'
                   '</p:clrMapOvr>'
                   '</p:notes>'
        )
        """From http://msdn.microsoft.com/en-us/library/office/gg278319%28v=office.15%29.aspx#sectionSection4
        """


            
        