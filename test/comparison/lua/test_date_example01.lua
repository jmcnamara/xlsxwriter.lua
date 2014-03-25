----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case for dates.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_date_example01.xlsx")
local worksheet = workbook:add_worksheet()

-- Widen column A for extra visibility.
worksheet:set_column("A:A", 30)

-- A number to convert to a date.
local number = 41333.5

-- Write it as a number without formatting.
worksheet:write("A1", number)          --> 41333.5

local format2 = workbook:add_format({num_format = "dd/mm/yy"})
worksheet:write("A2", number, format2) --> 28/02/13

local format3 = workbook:add_format({num_format = "mm/dd/yy"})
worksheet:write("A3", number, format3) --> 02/28/13

local format4 = workbook:add_format({num_format = "d\\-m\\-yyyy"})
worksheet:write("A4", number, format4) --> 28-2-2013

local format5 = workbook:add_format({num_format = "dd/mm/yy\\ hh:mm"})
worksheet:write("A5", number, format5) --> 28/02/13 12:00

local format6 = workbook:add_format({num_format = "d\\ mmm\\ yyyy"})
worksheet:write("A6", number, format6) --> 28 Feb 2013

local format7 = workbook:add_format({num_format = "mmm\\ d\\ yyyy\\ hh:mm\\ AM/PM"})
worksheet:write("A7", number, format7) --> Feb 28 2008 12:00 PM

workbook:close()
