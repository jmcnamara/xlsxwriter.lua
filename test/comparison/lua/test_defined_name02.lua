----
-- Test cases for xlsxwriter.lua.
--
-- Test defined names in the workbook.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook   = Workbook:new("test_defined_name02.xlsx")
local worksheet1 = workbook:add_worksheet('sheet One')

workbook:define_name("Sales", "='sheet One'!$G$1:$H$10")

workbook:close()
