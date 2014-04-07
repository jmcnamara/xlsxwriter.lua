----
-- Test cases for xlsxwriter.lua.
--
-- Test the worksheet set_start_page() method.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_set_start_page01.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:set_start_page(1)
worksheet:set_paper(9)

worksheet:write_string("A1", "Foo")

workbook:close()
