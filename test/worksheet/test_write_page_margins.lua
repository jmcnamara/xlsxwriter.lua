----
-- Tests for the xlsxwriter.lua Worksheet class.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"

plan(4)

----
-- Tests setup.
--
local expected
local got
local caption
local Worksheet = require "xlsxwriter.worksheet"
local worksheet

----
-- Test the _write_page_margins() method.
--
caption  = " \tWorksheet: _write_page_margins()"
expected = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:_write_page_margins(0.7, 0.7, 0.75, 0.75)

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_page_margins() method.
--
caption  = " \tWorksheet: _write_page_margins()"
expected = '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:set_margins(0.5, 0.5, 0.5, 0.5)

worksheet:_write_page_margins()

got = worksheet:_get_data()

is(got, expected, caption)

----
----
-- Test the _write_page_margins() method.
--
caption  = " \tWorksheet: _write_page_margins()"
expected = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.5" footer="0.3"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:set_header('', 0.5)

worksheet:_write_page_margins()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_page_margins() method.
--
caption  = " \tWorksheet: _write_page_margins()"
expected = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.5"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:set_footer('', 0.5)

worksheet:_write_page_margins()

got = worksheet:_get_data()

is(got, expected, caption)

