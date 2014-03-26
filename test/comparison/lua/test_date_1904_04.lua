----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case for dates.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_date_1904_04.xlsx", {date_1904 = true})
local worksheet = workbook:add_worksheet()
local format    = workbook:add_format{num_format = 14}

worksheet:set_column("A:A", 12)

worksheet:write_date_string("A1", "1904-01-01", format)
worksheet:write_date_string("A2", "1906-09-27", format)
worksheet:write_date_string("A3", "1917-09-09", format)
worksheet:write_date_string("A4", "1931-05-19", format)
worksheet:write_date_string("A5", "2177-10-15", format)
worksheet:write_date_string("A6", "4641-11-27", format)

workbook:close()
