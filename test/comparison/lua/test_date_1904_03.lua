----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case for dates.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_date_1904_03.xlsx")
local worksheet = workbook:add_worksheet()
local format    = workbook:add_format{num_format = 14}

worksheet:set_column("A:A", 12)

worksheet:write_date_string("A1", "1899-12-31", format)
worksheet:write_date_string("A2", "1902-09-26", format)
worksheet:write_date_string("A3", "1913-09-08", format)
worksheet:write_date_string("A4", "1927-05-18", format)
worksheet:write_date_string("A5", "2173-10-14", format)
worksheet:write_date_string("A6", "4637-11-26", format)

workbook:close()
