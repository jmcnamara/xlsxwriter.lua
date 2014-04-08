----
-- Test cases for xlsxwriter.lua.
--
-- Test merged_cells() method call.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_merge_cells02.xlsx")
local worksheet = workbook:add_worksheet()
local format    = workbook:add_format({align = "center"})

worksheet:merge_range("B1:B2", "col1", format)
worksheet:merge_range("C1:C2", "col2", format)
worksheet:merge_range("D1:D2", "col3", format)
worksheet:merge_range("E1:E2", "col4", format)

workbook:close()
