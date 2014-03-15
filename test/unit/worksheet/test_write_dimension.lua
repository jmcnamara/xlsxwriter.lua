----
-- Tests for the xlsxwriter.lua worksheet class.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"

plan(10)

----
-- Tests setup.
--
local expected
local got
local caption
local Worksheet     = require "xlsxwriter.worksheet"
local Sharedstrings = require "xlsxwriter.sharedstrings"
local worksheet
local cell_ref

----
-- 1. Test the _write_dimension() method with no dimensions set.
--
caption  = " \tWorksheet: _write_dimension()"
expected = '<dimension ref="A1"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:_write_dimension()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 2. Test the _write_dimension() method with dimensions set.
--
cell_ref = 'A1'
caption  = string.format(" \tWorksheet: _write_dimension('%s')", cell_ref)
expected = string.format('<dimension ref="%s"/>', cell_ref)

worksheet = Worksheet:new()
worksheet.str_table = Sharedstrings:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:write(cell_ref, 'some string')
worksheet:_write_dimension()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 3. Test the _write_dimension() method with dimensions set.
--
cell_ref = 'A1048576'
caption  = string.format(" \tWorksheet: _write_dimension('%s')", cell_ref)
expected = string.format('<dimension ref="%s"/>', cell_ref)

worksheet = Worksheet:new()
worksheet.str_table = Sharedstrings:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:write(cell_ref, 'some string')
worksheet:_write_dimension()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 4. Test the _write_dimension() method with dimensions set.
--
cell_ref = 'XFD1'
caption  = string.format(" \tWorksheet: _write_dimension('%s')", cell_ref)
expected = string.format('<dimension ref="%s"/>', cell_ref)

worksheet = Worksheet:new()
worksheet.str_table = Sharedstrings:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:write(cell_ref, 'some string')
worksheet:_write_dimension()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 5. Test the _write_dimension() method with dimensions set.
--
cell_ref = 'XFD1048576'
caption  = string.format(" \tWorksheet: _write_dimension('%s')", cell_ref)
expected = string.format('<dimension ref="%s"/>', cell_ref)

worksheet = Worksheet:new()
worksheet.str_table = Sharedstrings:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:write(cell_ref, 'some string')
worksheet:_write_dimension()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 6. Test the _write_dimension() method with dimensions set.
--
cell_ref = 'A1'
caption  = string.format(" \tWorksheet: _write_dimension('%s')", cell_ref)
expected = string.format('<dimension ref="%s"/>', cell_ref)

worksheet = Worksheet:new()
worksheet.str_table = Sharedstrings:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:write(cell_ref, 'some string')
worksheet:_write_dimension()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 7. Test the _write_dimension() method with dimensions set.
--
cell_ref = 'A1:B2'
caption  = string.format(" \tWorksheet: _write_dimension('%s')", cell_ref)
expected = string.format('<dimension ref="%s"/>', cell_ref)

worksheet = Worksheet:new()
worksheet.str_table = Sharedstrings:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:write('A1', 'some string')
worksheet:write('B2', 'some string')
worksheet:_write_dimension()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 8. Test the _write_dimension() method with dimensions set.
--
cell_ref = 'A1:B2'
caption  = string.format(" \tWorksheet: _write_dimension('%s')", cell_ref)
expected = string.format('<dimension ref="%s"/>', cell_ref)

worksheet = Worksheet:new()
worksheet.str_table = Sharedstrings:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:write('B2', 'some string')
worksheet:write('A1', 'some string')
worksheet:_write_dimension()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 9. Test the _write_dimension() method with dimensions set.
--
cell_ref = 'B2:H11'
caption  = string.format(" \tWorksheet: _write_dimension('%s')", cell_ref)
expected = string.format('<dimension ref="%s"/>', cell_ref)

worksheet = Worksheet:new()
worksheet.str_table = Sharedstrings:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:write('B2',  'some string')
worksheet:write('H11', 'some string')
worksheet:_write_dimension()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 10. Test the _write_dimension() method with dimensions set.
--
cell_ref = 'A1:XFD1048576'
caption  = string.format(" \tWorksheet: _write_dimension('%s')", cell_ref)
expected = string.format('<dimension ref="%s"/>', cell_ref)

worksheet = Worksheet:new()
worksheet.str_table = Sharedstrings:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:write('A1',         'some string')
worksheet:write('XFD1048576', 'some string')
worksheet:_write_dimension()

got = worksheet:_get_data()

is(got, expected, caption)

