----
-- Test cases for xlsxwriter.lua.
--
-- Test the creation of a simple xlsxwriter.lua file.
-- Check encoding of rich strings.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_escapes08.xlsx")
local worksheet = workbook:add_worksheet()

-- Test an already escapted string.
worksheet:write_url("A1", "http://example.com/%5b0%5d", nil,  "http://example.com/[0]")

workbook:close()

