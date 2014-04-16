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
local SST      = require "xlsxwriter.sharedstrings"

local workbook  = Workbook:new("test_hyperlink19.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:write_url("A1", "http://www.perl.com/")

-- Overwrite the URL string.
worksheet:write_formula("A1", "=1+1", nil, 2)

-- Reset the SST for testing.
workbook.str_table = SST:new()

workbook:close()

