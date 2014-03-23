----
-- Tests for the xlsxwriter.lua worksheet class.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"
require "Test.LongString"

plan(20)

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
  <cols>
    <col min="1" max="4" width="17.7109375" customWidth="1"/>
  </cols>
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

max_row = 1048576
max_col = 16384
bound_error = -1

-- Test some out of bound values.
got = worksheet:write_string(max_row, 0, 'Foo')
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_string(0, max_col, 'Foo')
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_string(max_row, max_col, 'Foo')
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_number(max_row, 0, 123)
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_number(0, max_col, 123)
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_number(max_row, max_col, 123)
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_blank(max_row, 0, nil, 'format')
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_blank(0, max_col, nil, 'format')
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_blank(max_row, max_col, nil, 'format')
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_formula(max_row, 0, '=A1')
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_formula(0, max_col, '=A1')
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_formula(max_row, max_col, '=A1')
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_array_formula(0, 0, 0, max_col, '=A1')
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_array_formula(0, 0, max_row, 0, '=A1')
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_array_formula(0, max_col, 0, 0, '=A1')
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_array_formula(max_row, 0, 0, 0, '=A1')
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:write_array_formula(max_row, max_col, max_row, max_col, '=A1')
is(got, bound_error, " \ttest write() bounds checks.")

-- Column out of bounds.
got = worksheet:set_column(6, max_col, 17)
is(got, bound_error, " \ttest write() bounds checks.")

got = worksheet:set_column(max_col, 6, 17)
is(got, bound_error, " \ttest write() bounds checks.")

-- Row out of bounds.
worksheet:set_row(max_row, 30)

-- Reverse man and min column numbers
worksheet:set_column(0, 3, 17)

-- Write some valid strings.
worksheet:write_string(0, 0, 'Foo')
worksheet:write_string(2, 0, 'Bar')
worksheet:write_string(2, 3, 'Baz')

worksheet:select()
worksheet:_assemble_xml_file()

got = _clean_xml_string(worksheet:_get_data())

is_string(got, expected, caption)
