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
local caption
local got
local Sharedstrings = require 'xlsxwriter.sharedstrings'
local sharedstrings

----
-- Test the _write_sst() method.
--
caption  = " \tSharedStrings: _write_sst()"
expected = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="7" uniqueCount="3">'

sharedstrings = Sharedstrings:new()
sharedstrings:_set_filehandle(io.tmpfile())

sharedstrings:_write_sst(7, 3)

got = sharedstrings:_get_data()

is(got, expected, caption)
