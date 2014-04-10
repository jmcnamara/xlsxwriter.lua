----
-- Tests for the xlsxwriter.lua.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"

plan(3)

----
-- Tests setup.
--
local expected
local got
local caption
local Worksheet = require 'xlsxwriter.worksheet'
local worksheet
local format = nil
local SST = require "xlsxwriter.sharedstrings"

----
-- Test the _write_merge_cells() method. With row, col notation.
--
caption  = " \tWorksheet: _write_merge_cells()"
expected = '<mergeCells count="1"><mergeCell ref="B3:C3"/></mergeCells>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet.str_table = SST:new()

worksheet:merge_range(2, 1, 2, 2, 'Foo', format)
worksheet:_write_merge_cells()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_merge_cells() method. With A1 notation.
--
caption  = " \tWorksheet: _write_merge_cells()"
expected = '<mergeCells count="1"><mergeCell ref="B3:C3"/></mergeCells>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet.str_table = SST:new()

worksheet:merge_range('B3:C3', 'Foo', format)
worksheet:_write_merge_cells()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_merge_cells() method. With more than one range.
--
caption  = " \tWorksheet: _write_merge_cells()"
expected = '<mergeCells count="2"><mergeCell ref="B3:C3"/><mergeCell ref="A2:D2"/></mergeCells>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet.str_table = SST:new()

worksheet:merge_range('B3:C3', 'Foo', format)
worksheet:merge_range('A2:D2', 'Foo', format)
worksheet:_write_merge_cells()

got = worksheet:_get_data()

is(got, expected, caption)

