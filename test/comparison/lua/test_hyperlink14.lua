----
-- Test cases for xlsxwriter.lua.
--
-- Test the creation of a simple xlsxwriter.lua file with hyperlinks.
-- This example has writes a url in a range.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_hyperlink14.xlsx")
local worksheet = workbook:add_worksheet()
local format    = workbook:add_format{["align"] = "center"}

worksheet:merge_range("C4:E5", "", format)
worksheet:write_url("C4", "http://www.perl.org/", format, "Perl Home")


workbook:close()

