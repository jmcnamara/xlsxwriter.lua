----
-- Tests for the xlsxwriter.lua.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"

plan(2)

----
-- Tests setup.
--
local expected
local got
local caption
local Styles = require 'xlsxwriter.styles'
local styles

----
-- 1. Test the _write_num_fmts() method.
--
caption  = " \tStyles: _write_num_fmts()"
expected = ""

styles = Styles:new()
styles:_set_filehandle(io.tmpfile())

styles:_write_num_fmts()

got = styles:_get_data()

is(got, expected, caption)

----
-- 2. Test the _write_num_fmts() method.
--
caption          = " \tStyles: _write_num_fmts()"
expected         = '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.0"/></numFmts>'

formats = {{["num_format_index"] = 164, ["num_format"] = '#,##0.0'}}

styles = Styles:new()
styles:_set_filehandle(io.tmpfile())
styles.num_format_count = 1
styles.xf_formats = formats

styles:_write_num_fmts()

got = styles:_get_data()

is(got, expected, caption)
