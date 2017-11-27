Comments
========

.. warning::

    Openpyxl currently supports the reading and writing of comment text only.
    Formatting information is lost.
    Comments are not currently supported if `use_iterators=True` is used.


Adding a comment to a cell
--------------------------

Comments have a text attribute and an author attribute, which must both be set

.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.comments import Comment
>>> wb = Workbook()
>>> ws = wb.active
>>> comment = ws["A1"].comment
>>> comment = Comment('This is the comment text', 'Comment Author')
>>> comment.text
'This is the comment text'
>>> comment.author
'Comment Author'

You cannot assign the same Comment object to two different cells. Doing so
raises an AttributeError.

.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.comments import Comment
>>> wb=Workbook()
>>> ws=wb.active
>>> comment = Comment("Text", "Author")
>>> ws["A1"].comment = comment
>>> ws["B2"].comment = comment
Traceback (most recent call last):
AttributeError: Comment already assigned to A1 in worksheet Sheet. Cannot
assign a comment to more than one cell


Loading and saving comments
----------------------------

Comments present in a workbook when loaded are stored in the comment
attribute of their respective cells automatically. Formatting information
such as font size, bold and italics are lost, as are the original dimensions
and position of the comment's container box.

Comments remaining in a workbook when it is saved are automatically saved to
the workbook file.
