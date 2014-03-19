----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case to test data writing.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_simple03.xlsx")

local worksheet1 = workbook:add_worksheet()
local worksheet2 = workbook:add_worksheet('Data Sheet')
local worksheet3 = workbook:add_worksheet()

local bold = workbook:add_format({bold = 1})

worksheet1:write('A1', 'Foo')
worksheet1:write('A2', 123)

worksheet3:write('B2', 'Foo')
worksheet3:write('B3', 'Bar', bold)
worksheet3:write('C4', 234)

worksheet2:activate()

worksheet2:select()
worksheet3:select()
worksheet3:activate()


workbook:close()
