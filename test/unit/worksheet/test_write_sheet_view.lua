----
-- Tests for the xlsxwriter.lua Worksheet class.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"

plan(6)

----
-- Tests setup.
--
local expected
local got
local caption
local Worksheet = require "xlsxwriter.worksheet"
local worksheet

----
-- Test the _write_sheet_view() method. Tab not selected.
--
caption  = " \tWorksheet: _write_sheet_view()"
expected = '<sheetView workbookViewId="0"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:_write_sheet_view()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_sheet_view() method. Tab selected.
--
caption  = " \tWorksheet: _write_sheet_view()"
expected = '<sheetView tabSelected="1" workbookViewId="0"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:select()
worksheet:_write_sheet_view()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_sheet_view() method. Tab selected + hide_gridlines().
--
caption  = " \tWorksheet: _write_sheet_view() + hide_gridlines()"
expected = '<sheetView tabSelected="1" workbookViewId="0"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:select()
worksheet:hide_gridlines()
worksheet:_write_sheet_view()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_sheet_view() method. Tab selected + hide_gridlines().
--
caption  = " \tWorksheet: _write_sheet_view() + hide_gridlines(0)"
expected = '<sheetView tabSelected="1" workbookViewId="0"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:select()
worksheet:hide_gridlines(0)
worksheet:_write_sheet_view()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_sheet_view() method. Tab selected + hide_gridlines().
--
caption  = " \tWorksheet: _write_sheet_view() + hide_gridlines(1)"
expected = '<sheetView tabSelected="1" workbookViewId="0"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:select()
worksheet:hide_gridlines(1)
worksheet:_write_sheet_view()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_sheet_view() method. Tab selected + hide_gridlines().
--
caption  = " \tWorksheet: _write_sheet_view() + hide_gridlines(2)"
expected = '<sheetView showGridLines="0" tabSelected="1" workbookViewId="0"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:select()
worksheet:hide_gridlines(2)
worksheet:_write_sheet_view()

got = worksheet:_get_data()

is(got, expected, caption)

