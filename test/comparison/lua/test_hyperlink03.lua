----
-- Test cases for xlsxwriter.lua.
--
-- Test the creation of a simple xlsxwriter.lua file with hyperlinks.
-- This example doesn't have any link formatting and tests the relationship
-- linkage code.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook   = Workbook:new("test_hyperlink03.xlsx")
local worksheet1 = workbook:add_worksheet()
local worksheet2 = workbook:add_worksheet()

worksheet1:write_url("A1",  "http://www.perl.org/")
worksheet1:write_url("D4",  "http://www.perl.org/")
worksheet1:write_url("A8",  "http://www.perl.org/")
worksheet1:write_url("B6",  "http://www.cpan.org/")
worksheet1:write_url("F12", "http://www.cpan.org/")

worksheet2:write_url("C2", "http://www.google.com/")
worksheet2:write_url("C5", "http://www.cpan.org/")
worksheet2:write_url("C7", "http://www.perl.org/")

workbook:close()

