----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case to test data writing.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_data03.xlsx")
local worksheet = workbook:add_worksheet()

-- Tests for range limits.
worksheet:write_number(0,       16383, 123)
worksheet:write_number(1048575, 16383, 456)

workbook:close()
