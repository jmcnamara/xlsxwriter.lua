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
local Format    = require "xlsxwriter.format"
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
  <dimension ref="A2:A15"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="2" ht="30" customHeight="1"/>
    <row r="4" hidden="1"/>
    <row r="7" s="1" customFormat="1"/>
    <row r="10" ht="3" customHeight="1"/>
    <row r="13" ht="24" hidden="1" customHeight="1"/>
    <row r="15" hidden="1"/>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>]])

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

local format = Format:new{xf_index = 1}

worksheet:set_row(1, 30)
worksheet:set_row(3, nil, nil, {['hidden'] = true})
worksheet:set_row(6, nil, format)
worksheet:set_row(9, 3)
worksheet:set_row(12, 24, nil, {['hidden'] = true})
worksheet:set_row(14, 0)

worksheet:select()
worksheet:_assemble_xml_file()

got = _clean_xml_string(worksheet:_get_data())

is_string(got, expected, caption)
