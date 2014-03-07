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
-- Test the _write_num_fmt() method.
--
caption  = " \tStyles: _write_num_fmt()"
expected = '<numFmt numFmtId="164" formatCode="#,##0.0"/>'

styles = Styles:new()
styles:_set_filehandle(io.tmpfile())

styles:_write_num_fmt(164, '#,##0.0')

got = styles:_get_data()

is(got, expected, caption)
