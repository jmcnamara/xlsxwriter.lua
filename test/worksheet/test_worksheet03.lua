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
  <dimension ref="F1:H1"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="2" max="4" width="5.7109375" customWidth="1"/>
    <col min="6" max="6" width="8.7109375" hidden="1" customWidth="1"/>
    <col min="8" max="8" width="9.140625" style="1"/>
    <col min="10" max="10" width="2.7109375" customWidth="1"/>
    <col min="12" max="12" width="0" hidden="1" customWidth="1"/>
  </cols>
  <sheetData/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>]])

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

local format = Format:new{xf_index = 1}

worksheet:set_column(1, 3, 5)
worksheet:set_column(5, 5, 8, nil, {['hidden'] = true})
worksheet:set_column(7, 7, nil, format)
worksheet:set_column(9, 9, 2)
worksheet:set_column(11, 11, nil, nil, {['hidden'] = true})

worksheet:select()
worksheet:_assemble_xml_file()

got = _clean_xml_string(worksheet:_get_data())

is_string(got, expected, caption)
