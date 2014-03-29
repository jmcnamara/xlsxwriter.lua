----
-- Test cases for xlsxwriter.lua.
--
-- Test the conversion of Lua types to Excel types.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_types12.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:write("A1", true)
worksheet:write("A2", false)

workbook:close()
