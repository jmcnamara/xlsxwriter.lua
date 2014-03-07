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
local Styles = require 'xlsxwriter.styles'
local styles

----
-- Test the xml_declaration() method.
--
caption  = " \tStyles: xml_declaration()"
expected = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

styles = Styles:new()
styles:_set_filehandle(io.tmpfile())

styles:_xml_declaration()

got = styles:_get_data()

is(got, expected, caption)

