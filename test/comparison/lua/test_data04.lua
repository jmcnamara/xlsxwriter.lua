----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case to test data writing.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_data04.xlsx")
local worksheet = workbook:add_worksheet()

-- Tests for the string table.
worksheet:write_string(      0, 0, "Foo")
worksheet:write_string(      0, 1, "Bar")
worksheet:write_string(      1, 0, "Bing")
worksheet:write_string(      2, 0, "Buzz")
worksheet:write_string(1048575, 0, "End")

workbook:close()
