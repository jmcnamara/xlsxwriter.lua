----
-- Test cases for xlsxwriter.lua.
--
-- Test the creation of a simple xlsxwriter.lua file with hyperlinks.
-- This example has link formatting.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook = Workbook:new("test_hyperlink20.xlsx")

-- Simulate custom colour for testing.
workbook.custom_colors = {"FF0000FF"}

local worksheet = workbook:add_worksheet()
local format1   = workbook:add_format{["font_color"] = "blue", ["underline"] = 1}
local format2   = workbook:add_format{["font_color"] = "red",  ["underline"] = 1}

worksheet:write_url("A1", "http://www.python.org/1", format1)
worksheet:write_url("A2", "http://www.python.org/2", format2)

workbook:close()

