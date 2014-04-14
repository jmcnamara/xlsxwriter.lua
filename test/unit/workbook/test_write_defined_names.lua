----
-- Tests for the xlsxwriter.lua.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"
require "Test.LongString"

plan(2)

----
-- Tests setup.
--
local expected
local got
local caption
local Workbook = require "xlsxwriter.workbook"
local workbook

----
-- Test the _write_defined_names() method.
--
caption  = " \tWorkbook: _write_defined_names()"
expected = '<definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!1:1</definedName></definedNames>'

workbook = Workbook:new("test")
workbook:_set_filehandle(io.tmpfile())
workbook.defined_names = {{'_xlnm.Print_Titles', 0, 'Sheet1!1:1'}}

workbook:_write_defined_names()

got = workbook:_get_data()

is(got, expected, caption)

----
-- Test the _write_defined_names() method.
--
caption  = " \tWorkbook: _write_defined_names()"
expected = [[<definedNames><definedName name="_Egg">Sheet1!A1</definedName><definedName name="_Fog">Sheet1!A1</definedName><definedName name="aaa" localSheetId="1">Sheet2!A1</definedName><definedName name="Abc">Sheet1!A1</definedName><definedName name="Bar" localSheetId="2">'Sheet 3'!A1</definedName><definedName name="Bar" localSheetId="0">Sheet1!A1</definedName><definedName name="Bar" localSheetId="1">Sheet2!A1</definedName><definedName name="Baz">0.98</definedName><definedName name="car" localSheetId="2">"Saab 900"</definedName></definedNames>]]

workbook = Workbook:new("test")
workbook:_set_filehandle(io.tmpfile())

workbook:add_worksheet()
workbook:add_worksheet()
workbook:add_worksheet('Sheet 3')

workbook:define_name("'Sheet 3'!Bar", "='Sheet 3'!A1")
workbook:define_name("Abc",           "=Sheet1!A1")
workbook:define_name("Baz",           "=0.98")
workbook:define_name("Sheet1!Bar",    "=Sheet1!A1")
workbook:define_name("Sheet2!Bar",    "=Sheet2!A1")
workbook:define_name("Sheet2!aaa",    "=Sheet2!A1")
workbook:define_name("'Sheet 3'!car", '="Saab 900"')
workbook:define_name("_Egg",          "=Sheet1!A1")
workbook:define_name("_Fog",          "=Sheet1!A1")

workbook:_prepare_defined_names()
workbook:_write_defined_names()

got = workbook:_get_data()

is_string(got, expected, caption)

