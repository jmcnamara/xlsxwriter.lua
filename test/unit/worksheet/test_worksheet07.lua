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
local SST = require "xlsxwriter.sharedstrings"

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
  <dimension ref="A1:C5"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="1" spans="1:3">
      <c r="A1">
        <v>1</v>
      </c>
    </row>
    <row r="2" spans="1:3">
      <c r="A2">
        <v>2</v>
      </c>
    </row>
    <row r="3" spans="1:3">
      <c r="C3">
        <f>A1+A2</f>
        <v>3</v>
      </c>
    </row>
    <row r="5" spans="1:3">
      <c r="B5" t="str">
        <f>"&lt;&amp;&gt;" &amp; ";"" '"</f>
        <v>&lt;&amp;&gt;;" '</v>
      </c>
    </row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>]])

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet.str_table = SST:new()

-- Write some data and formulas.
worksheet:write_number(0, 0, 1)
worksheet:write_number(1, 0, 2)
worksheet:write_formula(2, 2, '=A1+A2', nil, 3)
worksheet:write_formula(4, 1, [[="<&>" & ";"" '"]], nil, [[<&>;" ']])

worksheet:select()
worksheet:_assemble_xml_file()

got = _clean_xml_string(worksheet:_get_data())

is_string(got, expected, caption)
