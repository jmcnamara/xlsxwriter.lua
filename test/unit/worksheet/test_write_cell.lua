----
-- Tests for the xlsxwriter.lua Worksheet class.
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
local Worksheet = require "xlsxwriter.worksheet"
local worksheet
local format = nil

----
-- Test the _write_cell() method for numbers.
--
caption  = " \tWorksheet: _write_cell()"
expected = '<c r="A1"><v>1</v></c>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:_write_cell(0, 0, {'n', 1})

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_cell() method for strings.
--
caption  = " \tWorksheet: _write_cell()"
expected = '<c r="B4" t="s"><v>0</v></c>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:_write_cell(3, 1, {'s', 0})

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_cell() method for formulas with an optional value.
--
caption  = " \tWorksheet: _write_cell()"
expected = '<c r="C2"><f>A3+A5</f><v>0</v></c>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:_write_cell(1, 2, {'f', 'A3+A5', format, 0})

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_cell() method for formulas without an optional value.
--
caption  = " \tWorksheet: _write_cell()"
expected = '<c r="C2"><f>A3+A5</f><v>0</v></c>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:_write_cell(1, 2, {'f', 'A3+A5'})

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_cell() method for array formulas with an optional value.
--
caption  = " \tWorksheet: _write_cell()"
expected = '<c r="A1"><f t="array" ref="A1">SUM(B1:C1*B2:C2)</f><v>9500</v></c>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:_write_cell(0, 0, {'a', 'SUM(B1:C1*B2:C2)', format, 9500, 'A1'})

got = worksheet:_get_data()

is(got, expected, caption)

