----
-- Tests for the xlsxwriter.lua.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"

plan(8)

----
-- Tests setup.
--
local expected
local got
local caption
local Worksheet = require 'xlsxwriter.worksheet'
local worksheet

----
-- Test the _write_print_options() method. Without any options.
--
caption  = " \tWorksheet: _write_print_options()"
expected = ""

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:_write_print_options()

got = worksheet:_get_data()

is(got, expected, caption)
got = ""

----
-- Test the _write_print_options() method.
--
caption  = " \tWorksheet: _write_print_options()"
expected = '<printOptions horizontalCentered="1"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:center_horizontally()

worksheet:_write_print_options()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_print_options() method.
--
caption  = " \tWorksheet: _write_print_options()"
expected = '<printOptions verticalCentered="1"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:center_vertically()

worksheet:_write_print_options()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_print_options() method.
--
caption  = " \tWorksheet: _write_print_options()"
expected = '<printOptions horizontalCentered="1" verticalCentered="1"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:center_horizontally()
worksheet:center_vertically()

worksheet:_write_print_options()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_print_options() method.
--
caption  = " \tWorksheet: _write_print_options()"
expected = ''

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:hide_gridlines()

worksheet:_write_print_options()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_print_options() method.
--
caption  = " \tWorksheet: _write_print_options()"
expected = '<printOptions gridLines="1"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:hide_gridlines(0)

worksheet:_write_print_options()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_print_options() method.
--
caption  = " \tWorksheet: _write_print_options()"
expected = ''

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:hide_gridlines()

worksheet:_write_print_options(1)

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_print_options() method.
--
caption  = " \tWorksheet: _write_print_options()"
expected = '<printOptions gridLines="1"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:hide_gridlines(0)

worksheet:_write_print_options(2)

got = worksheet:_get_data()

is(got, expected, caption)

