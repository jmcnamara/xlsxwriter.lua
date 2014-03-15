----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case to test data writing.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_data06.xlsx")
local worksheet = workbook:add_worksheet()
local format1    = workbook:add_format{bold   = true}
local format2    = workbook:add_format{italic = true}
local format3    = workbook:add_format{bold   = true, italic = true}

-- Test some formatting.
worksheet:write('A1', "Foo", format1)
worksheet:write('A2', "Bar", format2)
worksheet:write('A3', "Baz", format3)

workbook:close()
