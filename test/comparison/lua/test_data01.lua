----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case to test data writing.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_data01.xlsx")
local worksheet = workbook:add_worksheet()

-- Write simple string value.
worksheet:write('A1', "Hello")

workbook:close()
