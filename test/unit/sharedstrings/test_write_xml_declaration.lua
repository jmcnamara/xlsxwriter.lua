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
-- Test the xml_declaration() method.
--
caption  = " \tSharedStrings: xml_declaration()"
expected = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

sharedstrings = Sharedstrings:new()
sharedstrings:_set_filehandle(io.tmpfile())

sharedstrings:_xml_declaration()

got = sharedstrings:_get_data()

is(got, expected, caption)
