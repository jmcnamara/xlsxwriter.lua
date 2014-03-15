----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case to test data writing.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_data05.xlsx")
local worksheet = workbook:add_worksheet()
local format    = workbook:add_format({bold = true})

-- Test the bold format.
worksheet:write('A1', "Foo", format)

workbook:close()
