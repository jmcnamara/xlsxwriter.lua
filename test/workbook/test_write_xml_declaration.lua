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
local Workbook = require 'xlsxwriter.workbook'
local workbook

----
-- Test the xml_declaration() method.
--
caption  = " \tWorkbook: xml_declaration()"
expected = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

workbook = Workbook:new("test.xlsx")
workbook:_set_filehandle(io.tmpfile())

workbook:_xml_declaration()

got = workbook:_get_data()

is(got, expected, caption)

