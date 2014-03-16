----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case for cell formatting.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_format05.xlsx")
local worksheet = workbook:add_worksheet()
local wrap      = workbook:add_format{text_wrap = true}

worksheet:set_row(0, 45)

worksheet:write("A1", "Foo\nBar", wrap)

workbook:close()
