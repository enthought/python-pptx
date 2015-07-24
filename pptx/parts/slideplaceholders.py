from pptx.shapes.factory import SlideShapeFactory
from pptx.shared import ParentedElementProxy

__author__ = 'ilewis'


class _SlidePlaceholders(ParentedElementProxy):
    """
    Collection of placeholder shapes on a slide. Supports iteration,
    :func:`len`, and dictionary-style lookup on the `idx` value of the
    placeholders it contains.
    """

    __slots__ = ()

    def __getitem__(self, idx):
        """
        Access placeholder shape having *idx*. Note that while this looks
        like list access, idx is actually a dictionary key and will raise
        |KeyError| if no placeholder with that idx value is in the
        collection.
        """
        for e in self._element.iter_ph_elms():
            if e.ph_idx == idx:
                return SlideShapeFactory(e, self)
        raise KeyError('no placeholder on this slide with idx == %d' % idx)

    def __iter__(self):
        """
        Generate placeholder shapes in `idx` order.
        """
        ph_elms = sorted(
            [e for e in self._element.iter_ph_elms()], key=lambda e: e.ph_idx
        )
        return (SlideShapeFactory(e, self) for e in ph_elms)

    def __len__(self):
        """
        Return count of placeholder shapes.
        """
        return len(list(self._element.iter_ph_elms()))