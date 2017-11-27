from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

from io import BytesIO
import datetime

from openpyxl.compat import unicode
from openpyxl.formatting import ConditionalFormatting

from openpyxl.xml.functions import SubElement
from openpyxl.styles import (
    Alignment,
    numbers,
    Color,
    Font,
    GradientFill,
    PatternFill,
    Border,
    Side,
    Protection,
    Style,
    colors,
    fills,
    borders,
)
from openpyxl.reader.excel import load_workbook

from openpyxl.writer.styles import StyleWriter
from openpyxl.writer.excel import save_virtual_workbook
from openpyxl.workbook import Workbook

from openpyxl.xml.functions import Element, tostring
from openpyxl.tests.helper import compare_xml


class DummyElement:

    def __init__(self):
        self.attrib = {}


class DummyWorkbook:

    style_properties = []
    _fonts = set()
    _borders = set()


def test_write_number_formats():
    wb = DummyWorkbook()
    wb._number_formats = ['YYYY']
    writer = StyleWriter(wb)
    writer._write_number_formats()
    xml = tostring(writer._root)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
       <numFmts count="1">
           <numFmt formatCode="YYYY" numFmtId="164"></numFmt>
        </numFmts>
    </styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


class TestStyleWriter(object):

    def setup(self):
        self.workbook = Workbook()
        self.worksheet = self.workbook.create_sheet()

    def test_no_style(self):
        w = StyleWriter(self.workbook)
        assert len(w.wb._cell_styles) == 1  # there is always the empty (defaul) style

    def test_nb_style(self):
        for i in range(1, 6):
            cell = self.worksheet.cell(row=1, column=i)
            cell.font = Font(size=i)
            _ = cell.style_id
        w = StyleWriter(self.workbook)
        assert len(w.wb._cell_styles) == 6  # 5 + the default

        cell = self.worksheet.cell('A10')
        cell.border=Border(top=Side(border_style=borders.BORDER_THIN))
        _ = cell.style_id
        w = StyleWriter(self.workbook)
        assert len(w.wb._cell_styles) == 7


    def test_default_xfs(self):
        w = StyleWriter(self.workbook)
        fonts = nft = borders = fills = DummyElement()
        w._write_cell_styles()
        xml = tostring(w._root)
        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <cellXfs count="1">
          <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
        </cellXfs>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_xfs_number_format(self):
        for idx, nf in enumerate(["0.0%", "0.00%", "0.000%"], 1):
            cell = self.worksheet.cell(row=idx, column=1)
            cell.number_format = nf
            _ = cell.style_id # add to workbook styles
        w = StyleWriter(self.workbook)
        w._write_cell_styles()

        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <cellXfs count="4">
                <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
                <xf borderId="0" fillId="0" fontId="0" numFmtId="164" xfId="0"/>
                <xf borderId="0" fillId="0" fontId="0" numFmtId="10" xfId="0"/>
                <xf borderId="0" fillId="0" fontId="0" numFmtId="165" xfId="0"/>
            </cellXfs>
        </styleSheet>
        """

        xml = tostring(w._root)
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_xfs_fonts(self):
        cell = self.worksheet.cell('A1')
        cell.font = Font(size=12, bold=True)
        _ = cell.style_id # update workbook styles
        w = StyleWriter(self.workbook)

        w._write_cell_styles()
        xml = tostring(w._root)

        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <cellXfs count="2">
            <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
            <xf borderId="0" fillId="0" fontId="1" numFmtId="0" xfId="0"/>
          </cellXfs>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_xfs_fills(self):
        cell = self.worksheet.cell('A1')
        cell.fill = fill=PatternFill(fill_type='solid',
                                     start_color=Color(colors.DARKYELLOW))
        _ = cell.style_id # update workbook styles
        w = StyleWriter(self.workbook)
        w._write_cell_styles()

        xml = tostring(w._root)
        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <cellXfs count="2">
            <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
            <xf borderId="0" fillId="2" fontId="0" numFmtId="0" xfId="0"/>
          </cellXfs>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_xfs_borders(self):
        cell = self.worksheet.cell('A1')
        cell.border=Border(top=Side(border_style=borders.BORDER_THIN,
                                    color=Color(colors.DARKYELLOW)))
        _ = cell.style_id # update workbook styles

        w = StyleWriter(self.workbook)
        w._write_cell_styles()

        xml = tostring(w._root)
        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <cellXfs count="2">
          <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
          <xf borderId="1" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
        </cellXfs>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_protection(self):
        cell = self.worksheet.cell('A1')
        cell.protection = Protection(locked=True, hidden=True)
        _ = cell.style_id

        w = StyleWriter(self.workbook)
        w._write_cell_styles()
        xml = tostring(w._root)
        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <cellXfs count="2">
            <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
            <xf applyProtection="1" borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0">
              <protection hidden="1" locked="1"/>
            </xf>
          </cellXfs>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_named_styles(self):
        writer = StyleWriter(self.workbook)
        writer._write_named_styles()
        xml = tostring(writer._root)
        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <cellStyleXfs count="1">
          <xf borderId="0" fillId="0" fontId="0" numFmtId="0"></xf>
        </cellStyleXfs>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_style_names(self):
        writer = StyleWriter(self.workbook)
        writer._write_style_names()
        xml = tostring(writer._root)
        expected = """
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <cellStyles count="1">
              <cellStyle name="Normal" xfId="0" builtinId="0" hidden="0"/>
            </cellStyles>
            </styleSheet>
            """
        diff = compare_xml(xml, expected)
        assert diff is None, diff



def test_simple_styles(datadir):
    wb = Workbook(guess_types=True)
    ws = wb.active
    now = datetime.datetime.now()
    for idx, v in enumerate(['12.34%', now, 'This is a test', '31.31415', None], 1):
        ws.append([v])
        _ = ws.cell(column=1, row=idx).style_id

    # set explicit formats
    ws['D9'].number_format = numbers.FORMAT_NUMBER_00
    ws['D9'].protection = Protection(locked=True)
    ws['D9'].style_id
    ws['E1'].protection = Protection(hidden=True)
    ws['E1'].style_id

    assert len(wb._cell_styles) == 5
    writer = StyleWriter(wb)

    datadir.chdir()
    with open('simple-styles.xml') as reference_file:
        expected = reference_file.read()
    xml = writer.write_table()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_empty_workbook():
    wb = Workbook()
    writer = StyleWriter(wb)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <numFmts count="0"/>
      <fonts count="1">
        <font>
          <name val="Calibri"/>
          <family val="2"/>
          <color theme="1"/>
          <sz val="11"/>
          <scheme val="minor"/>
        </font>
      </fonts>
      <fills count="2">
       <fill>
          <patternFill />
       </fill>
       <fill>
          <patternFill patternType="gray125"/>
        </fill>
      </fills>
      <borders count="1">
        <border>
          <left/>
          <right/>
          <top/>
          <bottom/>
          <diagonal/>
        </border>
      </borders>
      <cellStyleXfs count="1">
        <xf borderId="0" fillId="0" fontId="0" numFmtId="0"/>
      </cellStyleXfs>
      <cellXfs count="1">
        <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
      </cellXfs>
      <cellStyles count="1">
        <cellStyle builtinId="0" name="Normal" xfId="0" hidden="0"/>
      </cellStyles>
      <dxfs count="0"/>
      <tableStyles count="0" defaultPivotStyle="PivotStyleLight16" defaultTableStyle="TableStyleMedium9"/>
    </styleSheet>
    """
    xml = writer.write_table()
    diff = compare_xml(xml, expected)
    assert diff is None, diff
