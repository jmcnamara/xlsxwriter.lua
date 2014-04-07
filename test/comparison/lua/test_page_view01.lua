----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case for page view mode.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_page_view01.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:set_page_view()
worksheet:write_string("A1", "Foo")

workbook:close()
