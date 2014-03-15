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
-- Test the _write_styles_sheet() method.
--
caption  = " \tStyles: _write_styles_sheet()"
expected = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'

styles = Styles:new()
styles:_set_filehandle(io.tmpfile())

styles:_write_style_sheet()

got = styles:_get_data()

is(got, expected, caption)
