----
-- Test cases for xlsxwriter.lua.
--
-- Test the creation of a simple xlsxwriter.lua file.
-- Check encoding of rich strings.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_escapes04.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:write_url("A1", "http://www.perl.com/?a=1&b=2")

workbook:close()

