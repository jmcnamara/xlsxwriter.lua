----
-- Test cases for xlsxwriter.lua.
--
-- Test worksheet print options.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_print_options03.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:center_vertically()
worksheet:write_string("A1", "Foo")

workbook:close()
