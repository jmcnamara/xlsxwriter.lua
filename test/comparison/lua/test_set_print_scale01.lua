----
-- Test cases for xlsxwriter.lua.
--
-- Test the worksheet set_print_scale() method.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_set_print_scale01.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:set_print_scale(110)
worksheet:set_paper(9)

worksheet:write_string("A1", "Foo")

workbook:close()
