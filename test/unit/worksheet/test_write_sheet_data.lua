----
-- Tests for the xlsxwriter.lua Worksheet class.
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
-- Test the _write_sheet_data() method.
--
caption  = " \tWorksheet: _write_sheet_data()"
expected = '<sheetData/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:_write_sheet_data()

got = worksheet:_get_data()

is(got, expected, caption)

