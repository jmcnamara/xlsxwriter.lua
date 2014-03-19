----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case to test data writing.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_simple01.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:write_string("A1", "Hello")
worksheet:write_number("A2", 123)

workbook:close()
