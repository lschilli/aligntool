from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

# stdlib
import datetime
import decimal
from io import BytesIO

# package
from openpyxl import Workbook
from lxml.etree import xmlfile

# test imports
import pytest
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def worksheet():
    from openpyxl import Workbook
    wb = Workbook()
    return wb.active


@pytest.mark.lxml_required
@pytest.mark.parametrize("value, expected",
                         [
                             (9781231231230, """<c t="n" r="A1"><v>9781231231230</v></c>"""),
                             (decimal.Decimal('3.14'), """<c t="n" r="A1"><v>3.14</v></c>"""),
                             (1234567890, """<c t="n" r="A1"><v>1234567890</v></c>"""),
                             ("=sum(1+1)", """<c r="A1"><f>sum(1+1)</f><v></v></c>"""),
                             (True, """<c t="b" r="A1"><v>1</v></c>"""),
                             ("Hello", """<c t="s" r="A1"><v>0</v></c>"""),
                             ("", """<c r="A1" t="s"></c>"""),
                             (None, """<c r="A1" t="n"></c>"""),
                             (datetime.date(2011, 12, 25), """<c r="A1" t="n" s="1"><v>40902</v></c>"""),
                         ])
def test_write_cell(worksheet, value, expected):
    from .. lxml_worksheet import write_cell

    ws = worksheet
    cell = ws['A1']
    cell.value = value

    out = BytesIO()
    with xmlfile(out) as xf:
        cell = ws['A1']
        write_cell(xf, ws, cell, cell.has_style)
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.lxml_required
def test_write_cell_string(worksheet):
    from .. lxml_worksheet import write_cell

    ws = worksheet
    ws['A1'] = "Hello"

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, ws['A1'])
    assert ws.parent.shared_strings == ["Hello"]


@pytest.fixture
def write_rows():
    from .. lxml_worksheet import write_rows
    return write_rows


@pytest.mark.lxml_required
def test_write_sheetdata(worksheet, write_rows):
    ws = worksheet
    ws['A1'] = 10

    out = BytesIO()
    with xmlfile(out) as xf:
        write_rows(xf, ws)
    xml = out.getvalue()
    expected = """<sheetData><row r="1" spans="1:1"><c t="n" r="A1"><v>10</v></c></row></sheetData>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.lxml_required
def test_write_formula(worksheet, write_rows):
    ws = worksheet

    ws['F1'] = 10
    ws['F2'] = 32
    ws['F3'] = '=F1+F2'
    ws['A4'] = '=A1+A2+A3'
    ws['B4'] = "=SUM(A10:A14*B10:B14)"
    ws.formula_attributes['B4'] = {'t': 'array', 'ref': 'B4:B8'}

    out = BytesIO()
    with xmlfile(out) as xf:
        write_rows(xf, ws)

    xml = out.getvalue()
    expected = """
    <sheetData>
      <row r="1" spans="1:6">
        <c r="F1" t="n">
          <v>10</v>
        </c>
      </row>
      <row r="2" spans="1:6">
        <c r="F2" t="n">
          <v>32</v>
        </c>
      </row>
      <row r="3" spans="1:6">
        <c r="F3">
          <f>F1+F2</f>
          <v></v>
        </c>
      </row>
      <row r="4" spans="1:6">
        <c r="A4">
          <f>A1+A2+A3</f>
          <v></v>
        </c>
        <c r="B4">
          <f ref="B4:B8" t="array">SUM(A10:A14*B10:B14)</f>
          <v></v>
        </c>
      </row>
    </sheetData>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.lxml_required
def test_row_height(worksheet, write_rows):
    from openpyxl.worksheet.dimensions import RowDimension
    ws = worksheet
    ws['F1'] = 10
    ws.row_dimensions[1] = RowDimension(ws, height=30)
    ws.row_dimensions[2] = RowDimension(ws, height=30)

    out = BytesIO()
    with xmlfile(out) as xf:
        write_rows(xf, ws)
    xml = out.getvalue()
    expected = """
     <sheetData>
       <row customHeight="1" ht="30" r="1" spans="1:6">
         <c r="F1" t="n">
           <v>10</v>
         </c>
       </row>
       <row customHeight="1" ht="30" r="2" spans="1:6"></row>
     </sheetData>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
