----
-- Tests for the xlsxwriter.lua.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"

plan(5)

----
-- Tests setup.
--
local expected
local got
local caption
local Worksheet = require 'xlsxwriter.worksheet'
local worksheet

----
-- Test the _write_page_setup() method. Without any page setup.
--
caption  = " \tWorksheet: _write_page_setup()"
expected = ""

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:_write_page_setup()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_page_setup() method. With set_landscape()
--
caption  = " \tWorksheet: _write_page_setup()"
expected = '<pageSetup orientation="landscape"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:set_landscape()

worksheet:_write_page_setup()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_page_setup() method. With set_portrait()
--
caption  = " \tWorksheet: _write_page_setup()"
expected = '<pageSetup orientation="portrait"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:set_portrait()

worksheet:_write_page_setup()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_page_setup() method. With set_paper()
--
caption  = " \tWorksheet: _write_page_setup()"
expected = '<pageSetup paperSize="9" orientation="portrait"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:set_paper(9)

worksheet:_write_page_setup()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_page_setup() method. With print_across()
--
caption  = " \tWorksheet: _write_page_setup()"
expected = '<pageSetup pageOrder="overThenDown" orientation="portrait"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:print_across()

worksheet:_write_page_setup()

got = worksheet:_get_data()

is(got, expected, caption)

