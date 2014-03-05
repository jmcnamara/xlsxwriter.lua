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
-- 1. Test the _write_sheet_views() method.
--
caption  = " \tWorksheet: _write_sheet_views()"
expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:select()
worksheet:_write_sheet_views()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 2. Test the _write_sheet_views() method.
--
caption  = " \tWorksheet: _write_sheet_views()"
expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:select()
worksheet:set_zoom(100) -- Default. Should be ignored.
worksheet:_write_sheet_views()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 3. Test the _write_sheet_views() method. With zoom.
--
caption  = " \tWorksheet: _write_sheet_views()"
expected = '<sheetViews><sheetView tabSelected="1" zoomScale="200" zoomScaleNormal="200" workbookViewId="0"/></sheetViews>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:select()
worksheet:set_zoom(200)
worksheet:_write_sheet_views()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 4. Test the _write_sheet_views() method. Right to left.
--
caption  = " \tWorksheet: _write_sheet_views()"
expected = '<sheetViews><sheetView rightToLeft="1" tabSelected="1" workbookViewId="0"/></sheetViews>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:select()
worksheet:set_right_to_left()
worksheet:_write_sheet_views()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 5. Test the _write_sheet_views() method. Hide zeroes.
--
caption  = " \tWorksheet: _write_sheet_views()"
expected = '<sheetViews><sheetView showZeros="0" tabSelected="1" workbookViewId="0"/></sheetViews>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:select()
worksheet:hide_zero()
worksheet:_write_sheet_views()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 6. Test the _write_sheet_views() method. Set page view mode.
--
caption  = " \tWorksheet: _write_sheet_views()"
expected = '<sheetViews><sheetView tabSelected="1" view="pageLayout" workbookViewId="0"/></sheetViews>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:select()
worksheet:set_page_view()
worksheet:_write_sheet_views()

got = worksheet:_get_data()

is(got, expected, caption)

