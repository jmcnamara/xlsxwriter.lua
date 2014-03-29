----
-- Test cases for xlsxwriter.lua.
--
-- Test the conversion of Lua types to Excel types.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_types02.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:write_boolean("A1", true)
worksheet:write_boolean("A2", false)

workbook:close()
