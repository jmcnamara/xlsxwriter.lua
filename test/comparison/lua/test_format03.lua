----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case for cell formatting.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_format03.xlsx")
local worksheet = workbook:add_worksheet()

local format1 = workbook:add_format({bold   = 1, fg_color = "red"})
local format2 = workbook:add_format({bold   = 1, fg_color = "red", italic = 1})

worksheet:write("A1", "Foo" ,format1)
worksheet:write("A2", "Bar", format2)

workbook:close()
