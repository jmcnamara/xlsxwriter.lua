----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case for dates.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_date_1904_02.xlsx", {date_1904 = true})
local worksheet = workbook:add_worksheet()
local format    = workbook:add_format{num_format = 14}

worksheet:set_column("A:A", 12)

worksheet:write_date_time("A1", {year = 1904, month = 1,  day =  1}, format)
worksheet:write_date_time("A2", {year = 1906, month = 9,  day = 27}, format)
worksheet:write_date_time("A3", {year = 1917, month = 9,  day =  9}, format)
worksheet:write_date_time("A4", {year = 1931, month = 5,  day = 19}, format)
worksheet:write_date_time("A5", {year = 2177, month = 10, day = 15}, format)
worksheet:write_date_time("A6", {year = 4641, month = 11, day = 27}, format)

workbook:close()
