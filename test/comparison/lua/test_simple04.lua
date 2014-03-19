----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case to test data writing.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_simple04.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:set_column(0, 0, 12)

local format1 = workbook:add_format({num_format = 20})
local format2 = workbook:add_format({num_format = 14})

worksheet:write_date_string(0, 0, '12:00:00',   format1)
worksheet:write_date_string(1, 0, '2013-01-27', format2)

workbook:close()
