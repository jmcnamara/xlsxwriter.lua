----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case to test data writing.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_data07.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:write_formula("A1", "=1+2", false, 3)

workbook:close()
