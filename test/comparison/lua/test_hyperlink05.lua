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

local workbook  = Workbook:new("test_hyperlink05.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:write_url("A1", "http://www.perl.org/")
worksheet:write_url("A3", "http://www.perl.org/", nil, "Perl home")
worksheet:write_url("A5", "http://www.perl.org/", nil, "Perl home", "Tool Tip")
worksheet:write_url("A7", "http://www.cpan.org/", nil, "CPAN",      "Download")

workbook:close()
