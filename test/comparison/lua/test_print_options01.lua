----
-- Test cases for xlsxwriter.lua.
--
-- Test worksheet print options.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_print_options01.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:hide_gridlines(0)
worksheet:write_string("A1", "Foo")

workbook:close()
