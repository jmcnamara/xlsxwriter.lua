----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case to test data writing.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_data02.xlsx")
local worksheet = workbook:add_worksheet()

-- Tests for the limits of the row range.
worksheet:write(0,       0, 123)
worksheet:write(1048575, 0, 456)

-- These should be ignored.
worksheet:write_number(-1,      0, 123)
worksheet:write_number(1048576, 0, 456)

workbook:close()
