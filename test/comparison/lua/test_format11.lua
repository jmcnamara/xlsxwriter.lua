----
-- Test cases for xlsxwriter.lua.
--
-- Test a vertical and horizontal centered format.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_format11.xlsx")
local worksheet = workbook:add_worksheet()

local centered = workbook:add_format({align  = "center", valign = "vcenter"})

worksheet:write("B2", "Foo", centered)

workbook:close()

