----
-- Test cases for xlsxwriter.lua.
--
-- Test worksheet print options.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_print_options06.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:print_area("A1:G20")
worksheet:repeat_rows(0)

worksheet:write_string("A1", "Foo")

workbook:close()
