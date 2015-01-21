----
-- Tests for the xlsxwriter.lua worksheet class.
--
-- Copyright 2014-2015, John McNamara, jmcnamara@cpan.org
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
  <dimension ref="A1:D3"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="1" spans="1:4">
      <c r="A1" t="s">
        <v>0</v>
      </c>
    </row>
    <row r="3" spans="1:4">
      <c r="A3" t="s">
        <v>1</v>
      </c>
      <c r="D3" t="s">
        <v>2</v>
      </c>
    </row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>]])

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet.str_table = SST:new()

worksheet:write_string(0, 0, 'Foo')
worksheet:write_string(2, 0, 'Bar')
worksheet:write_string(2, 3, 'Baz')

worksheet:select()
worksheet:_assemble_xml_file()

got = _clean_xml_string(worksheet:_get_data())

is_string(got, expected, caption)
