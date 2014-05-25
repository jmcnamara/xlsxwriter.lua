----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case for cell formatting.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_format06.xlsx")
local worksheet = workbook:add_worksheet()

local format1 = workbook:add_format({num_format = 2})
local format2 = workbook:add_format({num_format = 12})

worksheet:write(0, 0, 1.2222)
worksheet:write(1, 0, 1.2222, format1)
worksheet:write(2, 0, 1.2222, format2)
worksheet:write(3, 0, 1.2222)
worksheet:write(4, 0, 1.2222)

workbook:close()
