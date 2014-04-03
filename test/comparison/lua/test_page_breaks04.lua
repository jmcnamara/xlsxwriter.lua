----
-- Test cases for xlsxwriter.lua.
--
-- Test the worksheet set_x_pagebreaks() methods.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_page_breaks04.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:set_v_pagebreaks({1})

worksheet:write_string("A1", "Foo")

workbook:close()
