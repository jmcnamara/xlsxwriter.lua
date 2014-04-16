----
-- Test cases for xlsxwriter.lua.
--
-- Test the creation of a simple xlsxwriter.lua file.
-- Check encoding of rich strings.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook   = Workbook:new("test_escapes05.xlsx")
local worksheet1 = workbook:add_worksheet("Start")
local worksheet2 = workbook:add_worksheet("A & B")

worksheet1:write_url("A1", [[internal:'A & B'!A1]], nil, "Jump to A & B")

workbook:close()

