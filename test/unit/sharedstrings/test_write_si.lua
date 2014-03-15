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
local SharedStrings = require 'xlsxwriter.sharedstrings'
local sharedstrings

----
-- Test the _write_si() method.
--
caption  = " \tSharedStrings: _write_si()"
expected = '<si><t>neptune</t></si>'

sharedstrings = SharedStrings:new()
sharedstrings:_set_filehandle(io.tmpfile())

sharedstrings:_write_si('neptune')

got = sharedstrings:_get_data()

is(got, expected, caption)
