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

local workbook  = Workbook:new("test_hyperlink02.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:write_url("A1",  "http://www.perl.org/")
worksheet:write_url("D4",  "http://www.perl.org/")
worksheet:write_url("A8",  "http://www.perl.org/")
worksheet:write_url("B6",  "http://www.cpan.org/")
worksheet:write_url("F12", "http://www.cpan.org/")

workbook:close()

