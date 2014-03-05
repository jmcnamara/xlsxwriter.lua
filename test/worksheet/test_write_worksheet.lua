----
-- Tests for the xlsxwriter.lua worksheet class.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"

plan(1)

----
-- Tests setup.
--
local expected
local got
local caption
local Worksheet = require "xlsxwriter.worksheet"
local worksheet

----
-- Test the _write_worksheet() method.
--
caption = " \tWorksheet: _write_worksheet()"
expected =
'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:_write_worksheet()
got = worksheet:_get_data()

is(got, expected, caption)
