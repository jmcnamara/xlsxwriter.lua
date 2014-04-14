----
-- Tests for the xlsxwriter.lua.
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
local Worksheet = require 'xlsxwriter.worksheet'
local worksheet

----
-- Test the _write_hyperlink_external() method.
--
caption  = " \tWorksheet: _write_hyperlink_external()"
expected = '<hyperlink ref="A1" r:id="rId1"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:_write_hyperlink_external(0, 0, 1)

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_hyperlink_internal() method.
--
caption  = " \tWorksheet: _write_hyperlink_internal()"
expected = '<hyperlink ref="A1" location="Sheet2!A1" display="Sheet2!A1"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:_write_hyperlink_internal(0, 0, 'Sheet2!A1', 'Sheet2!A1')

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_hyperlink_internal() method.
--
caption  = " \tWorksheet: _write_hyperlink_internal()"
expected = [[<hyperlink ref="A5" location="'Data Sheet'!D5" display="'Data Sheet'!D5"/>]]

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:_write_hyperlink_internal(4, 0, "'Data Sheet'!D5", "'Data Sheet'!D5")

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_hyperlink_internal() method.
--
caption  = " \tWorksheet: _write_hyperlink_internal()"
expected = '<hyperlink ref="A18" location="Sheet2!A1" tooltip="Screen Tip 1" display="Sheet2!A1"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:_write_hyperlink_internal(17, 0, 'Sheet2!A1', 'Sheet2!A1', 'Screen Tip 1')

got = worksheet:_get_data()

is(got, expected, caption)

