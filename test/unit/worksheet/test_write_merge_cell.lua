----
-- Tests for the xlsxwriter.lua.
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
local Worksheet = require 'xlsxwriter.worksheet'
local worksheet

----
-- Test the _write_merge_cell() method.
--
caption  = " \tWorksheet: _write_merge_cell()"
expected = '<mergeCell ref="B3:C3"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:_write_merge_cell({2, 1, 2, 2})

got = worksheet:_get_data()

is(got, expected, caption)
