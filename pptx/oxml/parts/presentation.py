# encoding: utf-8

"""
Custom element classes for presentation-related XML elements.
"""

from __future__ import absolute_import

from ..simpletypes import ST_SlideId, ST_SlideSizeCoordinate, XsdString
from ..xmlchemy import (
    BaseOxmlElement, RequiredAttribute, ZeroOrOne, ZeroOrMore
)


class CT_Presentation(BaseOxmlElement):
    """
    ``<p:presentation>`` element, root of the Presentation part stored as
    ``/ppt/presentation.xml``.
    """
    sldMasterIdLst = ZeroOrOne('p:sldMasterIdLst', successors=(
        'p:notesMasterIdLst', 'p:handoutMasterIdLst', 'p:sldIdLst',
        'p:sldSz', 'p:notesSz'
    ))
    notesMasterIdLst = ZeroOrOne('p:notesMasterIdLst', successors=(
        'p:handoutMasterIdLst', 'p:sldIdLst',
        'p:sldSz', 'p:notesSz'))
    sldIdLst = ZeroOrOne('p:sldIdLst', successors=('p:sldSz', 'p:notesSz'))
    sldSz = ZeroOrOne('p:sldSz', successors=('p:notesSz',))


class CT_SlideId(BaseOxmlElement):
    """
    ``<p:sldId>`` element, direct child of <p:sldIdLst> that contains an rId
    reference to a slide in the presentation.
    """
    id = RequiredAttribute('id', ST_SlideId)
    rId = RequiredAttribute('r:id', XsdString)


class CT_SlideIdList(BaseOxmlElement):
    """
    ``<p:sldIdLst>`` element, direct child of <p:presentation> that contains
    a list of the slide parts in the presentation.
    """
    sldId = ZeroOrMore('p:sldId')

    def add_sldId(self, rId):
        """
        Return a reference to a newly created <p:sldId> child element having
        its r:id attribute set to *rId*.
        """
        return self._add_sldId(id=self._next_id, rId=rId)

    @property
    def _next_id(self):
        """
        Return the next available slide ID as an int. Valid slide IDs start
        at 256. Unused ids in the sequences starting from 256 are used first.
        """
        id_str_lst = self.xpath('./p:sldId/@id')
        used_ids = [int(id_str) for id_str in id_str_lst]
        for n in range(256, 258+len(used_ids)):
            if n not in used_ids:
                return n


class CT_SlideMasterIdList(BaseOxmlElement):
    """
    ``<p:sldMasterIdLst>`` element, child of ``<p:presentation>`` containing
    references to the slide masters that belong to the presentation.
    """
    sldMasterId = ZeroOrMore('p:sldMasterId')



class CT_SlideMasterIdListEntry(BaseOxmlElement):
    """
    ``<p:sldMasterId>`` element, child of ``<p:sldMasterIdLst>`` containing
    a reference to a slide master.
    """
    rId = RequiredAttribute('r:id', XsdString)


class CT_NotesMasterIdList(BaseOxmlElement):
    """
    ``<p:sldMasterIdLst>`` element, child of ``<p:presentation>`` containing
    references to the slide masters that belong to the presentation.
    """
    notesMasterId = ZeroOrMore('p:notesMasterId')



class CT_NotesMasterIdListEntry(BaseOxmlElement):
    """
    ``<p:sldMasterId>`` element, child of ``<p:sldMasterIdLst>`` containing
    a reference to a slide master.
    """
    rId = RequiredAttribute('r:id', XsdString)


class CT_SlideSize(BaseOxmlElement):
    """
    ``<p:sldSz>`` element, direct child of <p:presentation> that contains the
    width and height of slides in the presentation.
    """
    cx = RequiredAttribute('cx', ST_SlideSizeCoordinate)
    cy = RequiredAttribute('cy', ST_SlideSizeCoordinate)
