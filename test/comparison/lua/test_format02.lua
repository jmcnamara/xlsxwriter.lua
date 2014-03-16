----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case for cell formatting.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_format02.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:set_row(0, 30)

local format1 = workbook:add_format({
    font_name = "Arial",
    bold      = 1,
    locked    = 1,
    align     = "left",
    valign    = "bottom"
})

local format2 = workbook:add_format({
    font_name = "Arial",
    bold      = 1,
    locked    = 1,
    rotation  = 90,
    align     = "center",
    valign    = "bottom"
})

worksheet:write("A1", "Foo", format1)
worksheet:write("B1", "Bar", format2)

workbook:close()
