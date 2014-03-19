----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case to test data writing.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_simple05.xlsx")
local worksheet = workbook:add_worksheet()

local format1 = workbook:add_format({bold           = 1})
local format2 = workbook:add_format({italic         = 1})
local format3 = workbook:add_format({bold           = 1, italic = 1})
local format4 = workbook:add_format({underline      = 1})
local format5 = workbook:add_format({font_strikeout = 1})
local format6 = workbook:add_format({font_script    = 1})
local format7 = workbook:add_format({font_script    = 2})

worksheet:set_row(5, 18)
worksheet:set_row(6, 18)

worksheet:write_string(0, 0, 'Foo', format1)
worksheet:write_string(1, 0, 'Foo', format2)
worksheet:write_string(2, 0, 'Foo', format3)
worksheet:write_string(3, 0, 'Foo', format4)
worksheet:write_string(4, 0, 'Foo', format5)
worksheet:write_string(5, 0, 'Foo', format6)
worksheet:write_string(6, 0, 'Foo', format7)

workbook:close()
