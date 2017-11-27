Manipulating a workbook in memory
=================================

Create a workbook
-----------------

There is no need to create a file on the filesystem to get started with openpyxl.
Just import the Workbook class and start using it ::

    >>> from openpyxl import Workbook
    >>> wb = Workbook()

A workbook is always created with at least one worksheet. You can get it by
using the :func:`openpyxl.workbook.Workbook.active` property ::

    >>> ws = wb.active

.. note::

    This function uses the `_active_sheet_index` property, set to 0 by default.
    Unless you modify its value, you will always get the
    first worksheet by using this method.

You can also create new worksheets by using the
:func:`openpyxl.workbook.Workbook.create_sheet` method ::

    >>> ws1 = wb.create_sheet() # insert at the end (default)
    # or
    >>> ws2 = wb.create_sheet(0) # insert at first position

Sheets are given a name automatically when they are created.
They are numbered in sequence (Sheet, Sheet1, Sheet2, ...).
You can change this name at any time with the `title` property::

    ws.title = "New Title"

The background color of the tab holding this title is white by default.
You can change this providing an RRGGBB color code to the sheet_properties.tabColor property::

    ws.sheet_properties.tabColor = "1072BA"

Once you gave a worksheet a name, you can get it as a key of the workbook or
using the :func:`openpyxl.workbook.Workbook.get_sheet_by_name` method ::

    >>> ws3 = wb["New Title"]
    >>> ws4 = wb.get_sheet_by_name("New Title")
    >>> ws is ws3 is ws4
    True

You can review the names of all worksheets of the workbook with the
:func:`openpyxl.workbook.Workbook.get_sheet_names` method ::

    >>> print(wb.get_sheet_names())
    ['Sheet2', 'New Title', 'Sheet1']

You can loop through worksheets ::

    >>> for sheet in wb:
    ...     print(sheet.title)


Playing with data
------------------

Accessing one cell
++++++++++++++++++

Now we know how to access a worksheet, we can start modifying cells content.

Cells can be accessed directly as keys of the worksheet ::

    >>> c = ws['A4']

This will return the cell at A4 or create one if it does not exist yet.
Values can be directly assigned ::

    >>> ws['A4'] = 4

There is also the :func:`openpyxl.worksheet.Worksheet.cell` method::

    >>> c = ws.cell('A4')

You can also access a cell using row and column notation::

    >>> d = ws.cell(row = 4, column = 2)

.. note::

    When a worksheet is created in memory, it contains no `cells`. They are
    created when first accessed. This way we don't create objects that would never
    be accessed, thus reducing the memory footprint.

.. warning::

    Because of this feature, scrolling through cells instead of accessing them
    directly will create them all in memory, even if you don't assign them a value.

    Something like ::

        >>> for i in range(1,101):
        ...        for j in range(1,101):
        ...            ws.cell(row = i, column = j)

    will create 100x100 cells in memory, for nothing.

    However, there is a way to clean all those unwanted cells, we'll see that later.


Accessing many cells
++++++++++++++++++++

Ranges of cells can be accessed using slicing ::

    >>> cell_range = ws['A1':'C2']

You can also use the :func:`openpyxl.worksheet.Worksheet.iter_rows` method::

    >>> tuple(ws.iter_rows('A1:C2'))
    ((<Cell Sheet1.A1>, <Cell Sheet1.B1>, <Cell Sheet1.C1>),
     (<Cell Sheet1.A2>, <Cell Sheet1.B2>, <Cell Sheet1.C2>))

    >>> for row in ws.iter_rows('A1:C2'):
    ...        for cell in row:
    ...            print cell
    <Cell Sheet1.A1>
    <Cell Sheet1.B1>
    <Cell Sheet1.C1>
    <Cell Sheet1.A2>
    <Cell Sheet1.B2>
    <Cell Sheet1.C2>

If you need to iterate through all the rows or columns of a file, you can instead use the
:func:`openpyxl.worksheet.Worksheet.rows` property::

    >>> ws = wb.active
    >>> ws['C9'] = 'hello world'
    >>> ws.rows
    ((<Cell Sheet.A1>, <Cell Sheet.B1>, <Cell Sheet.C1>),
    (<Cell Sheet.A2>, <Cell Sheet.B2>, <Cell Sheet.C2>),
    (<Cell Sheet.A3>, <Cell Sheet.B3>, <Cell Sheet.C3>),
    (<Cell Sheet.A4>, <Cell Sheet.B4>, <Cell Sheet.C4>),
    (<Cell Sheet.A5>, <Cell Sheet.B5>, <Cell Sheet.C5>),
    (<Cell Sheet.A6>, <Cell Sheet.B6>, <Cell Sheet.C6>),
    (<Cell Sheet.A7>, <Cell Sheet.B7>, <Cell Sheet.C7>),
    (<Cell Sheet.A8>, <Cell Sheet.B8>, <Cell Sheet.C8>),
    (<Cell Sheet.A9>, <Cell Sheet.B9>, <Cell Sheet.C9>))

or the :func:`openpyxl.worksheet.Worksheet.columns` property::

    >>> ws.columns
    ((<Cell Sheet.A1>,
    <Cell Sheet.A2>,
    <Cell Sheet.A3>,
    <Cell Sheet.A4>,
    <Cell Sheet.A5>,
    <Cell Sheet.A6>,
    ...
    <Cell Sheet.B7>,
    <Cell Sheet.B8>,
    <Cell Sheet.B9>),
    (<Cell Sheet.C1>,
    <Cell Sheet.C2>,
    <Cell Sheet.C3>,
    <Cell Sheet.C4>,
    <Cell Sheet.C5>,
    <Cell Sheet.C6>,
    <Cell Sheet.C7>,
    <Cell Sheet.C8>,
    <Cell Sheet.C9>))


Data storage
++++++++++++

Once we have a :class:`openpyxl.cell.Cell`, we can assign it a value::

    >>> c.value = 'hello, world'
    >>> print(c.value)
    'hello, world'

    >>> d.value = 3.14
    >>> print(d.value)
    3.14

You can also enable type and format inference::

    >>> wb = Workbook(guess_types=True)
    >>> c.value = '12%'
    >>> print(c.value)
    0.12

    >>> import datetime
    >>> d.value = datetime.datetime.now()
    >>> print d.value
    datetime.datetime(2010, 9, 10, 22, 25, 18)

    >>> c.value = '31.50'
    >>> print(c.value)
    31.5


Saving to a file
================

The simplest and safest way to save a workbook is by using the
:func:`openpyxl.workbook.Workbook.save()` method of the
:class:`openpyxl.workbook.Workbook` object::

    >>> wb = Workbook()
    >>> wb.save('balances.xlsx')

.. warning::

   This operation will overwrite existing files without warning.

.. note::

    Extension is not forced to be xlsx or xlsm, although you might have
    some trouble opening it directly with another application if you don't
    use an official extension.

    As OOXML files are basically ZIP files, you can also end the filename
    with .zip and open it with your favourite ZIP archive manager.

You can specify the attribute as_template=True, to save the document
as a template

    >>> wb = load_workbook('document.xlsx')# doctest: +SKIP
    >>> wb.save('document_template.xltx', as_template=True)# doctest: +SKIP

or specify the attribute as_template=False (by default), to save
the document template (or document) as document.

    >>> wb = load_workbook('document_template.xltx')# doctest: +SKIP
    >>> wb.save('document.xlsx', as_template=False)# doctest: +SKIP

    >>> wb = load_workbook('document.xlsx')# doctest: +SKIP
    >>> wb.save('new_document.xlsx', as_template=False)# doctest: +SKIP

.. warning::

    You should monitor the data attributes and document extensions
    for saving documents in the document templates and vice versa,
    otherwise the result table engine can not open the document.

.. note::

    The following will fail:

    >>> wb = load_workbook('document.xlsx')# doctest: +SKIP
    >>> # Need to save with the extension *.xlsx
    >>> wb.save('new_document.xlsm')# doctest: +SKIP
    >>> # MS Excel can't open the document
    >>>
    >>> # or
    >>>
    >>> # Need specify attribute keep_vba=True
    >>> wb = load_workbook('document.xlsm')# doctest: +SKIP
    >>> wb.save('new_document.xlsm')# doctest: +SKIP
    >>> # MS Excel can't open the document
    >>>
    >>> # or
    >>>
    >>> wb = load_workbook('document.xltm', keep_vba=True)# doctest: +SKIP
    >>> # If us need template document, then we need specify extension as *.xltm.
    >>> # If us need document, then we need specify attribute as_template=False.
    >>> wb.save('new_document.xlsm', as_template=True)# doctest: +SKIP
    >>> # MS Excel can't open the document


Loading from a file
===================

The same way as writing, you can import :func:`openpyxl.load_workbook` to
open an existing workbook::

    >>> from openpyxl import load_workbook
    >>> wb2 = load_workbook('test.xlsx')
    >>> print wb2.get_sheet_names()
    ['Sheet2', 'New Title', 'Sheet1']

This ends the tutorial for now, you can proceed to the :doc:`usage` section
