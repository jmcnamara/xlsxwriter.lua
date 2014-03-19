----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case to test data writing.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_simple02.xlsx")

worksheet1 = workbook:add_worksheet()
workbook:add_worksheet('Data Sheet')
worksheet3 = workbook:add_worksheet()

bold = workbook:add_format({bold = 1})

worksheet1:write_string(0, 0, 'Foo')
worksheet1:write_number(1, 0, 123)

worksheet3:write_string(1, 1, 'Foo')
worksheet3:write_string(2, 1, 'Bar', bold)
worksheet3:write_number(3, 2, 234)


workbook:close()
