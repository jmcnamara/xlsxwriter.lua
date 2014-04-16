----
-- Test cases for xlsxwriter.lua.
--
-- Test the creation of a simple xlsxwriter.lua file with hyperlinks.
-- This example has link formatting.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_hyperlink12.xlsx")
local worksheet = workbook:add_worksheet()
local format    = workbook:add_format{["font_color"] = "blue", ["underline"] = "1"}

worksheet:write_url("A1", "mailto:jmcnamara@cpan.org", format)

worksheet:write_url("A3", "ftp://perl.org/", format)

workbook:close()

