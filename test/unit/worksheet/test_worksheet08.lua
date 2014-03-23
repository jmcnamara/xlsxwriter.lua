----
-- Tests for the xlsxwriter.lua worksheet class.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"
require "Test.LongString"

plan(1)

----
-- Tests setup.
--
local expected
local got
local caption
local Worksheet = require "xlsxwriter.worksheet"
local worksheet

-- Remove extra whitespace in the formatted XML strings.
function _clean_xml_string(s)
  return (string.gsub(s, ">%s+<", "><"))
end

----
-- Test the _write_worksheet() method.
--
caption = " \tWorksheet: Worksheet: _assemble_xml_file()"
expected = _clean_xml_string([[
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:C3"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="1" spans="1:3">
      <c r="A1">
        <f t="array" ref="A1:A3">SUM(B1:C1*B2:C2)</f>
        <v>0</v>
      </c>
      <c r="B1">
        <v>0</v>
      </c>
      <c r="C1">
        <v>0</v>
      </c>
    </row>
    <row r="2" spans="1:3">
      <c r="A2">
        <v>0</v>
      </c>
      <c r="B2">
        <v>0</v>
      </c>
      <c r="C2">
        <v>0</v>
      </c>
    </row>
    <row r="3" spans="1:3">
      <c r="A3">
        <v>0</v>
      </c>
      <c r="B3">
        <v>0</v>
      </c>
      <c r="C3">
        <v>0</v>
      </c>
    </row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>]])

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:write_array_formula(0, 0, 2, 0, '{=SUM(B1:C1*B2:C2)}')
worksheet:write_number(0, 1, 0)
worksheet:write_number(1, 1, 0)
worksheet:write_number(2, 1, 0)
worksheet:write_number(0, 2, 0)
worksheet:write_number(1, 2, 0)
worksheet:write_number(2, 2, 0)

worksheet:select()
worksheet:_assemble_xml_file()

got = _clean_xml_string(worksheet:_get_data())

is_string(got, expected, caption)
