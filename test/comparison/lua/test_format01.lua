----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case for cell formatting.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook   = Workbook:new("test_format01.xlsx")
local worksheet1 = workbook:add_worksheet()
local worksheet2 = workbook:add_worksheet("Data Sheet")
local worksheet3 = workbook:add_worksheet()

local unused1 = workbook:add_format({bold   = 1})
local bold    = workbook:add_format({bold   = 1})
local unused2 = workbook:add_format({bold   = 1})
local unused3 = workbook:add_format({italic = 1})

worksheet1:write("A1", "Foo")
worksheet1:write("A2", 123)

worksheet3:write("B2", "Foo")
worksheet3:write("B3", "Bar", bold)
worksheet3:write("C4", 234)

workbook:close()
