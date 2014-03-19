----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case for dates.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_date_1904_01.xlsx")
local worksheet = workbook:add_worksheet()
local format    = workbook:add_format{num_format = 14}

worksheet:set_column("A:A", 12)

worksheet:write_date_time("A1", {year = 1899, month = 12, day = 31}, format)
worksheet:write_date_time("A2", {year = 1902, month = 9,  day = 26}, format)
worksheet:write_date_time("A3", {year = 1913, month = 9,  day =  8}, format)
worksheet:write_date_time("A4", {year = 1927, month = 5,  day = 18}, format)
worksheet:write_date_time("A5", {year = 2173, month = 10, day = 14}, format)
worksheet:write_date_time("A6", {year = 4637, month = 11, day = 26}, format)

workbook:close()
