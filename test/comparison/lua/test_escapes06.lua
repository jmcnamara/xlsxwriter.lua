----
-- Test cases for xlsxwriter.lua.
--
-- Test the creation of a simple xlsxwriter.lua file a num format that
-- require XML escaping.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook   = Workbook:new("test_escapes06.xlsx")
local worksheet  = workbook:add_worksheet()
local num_format = workbook:add_format{["num_format"] = '[Red]0.0%\\ "a"'}

worksheet:set_column("A:A", 14)

worksheet:write("A1", 123, num_format)

workbook:close()

