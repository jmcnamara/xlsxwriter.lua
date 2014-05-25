----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case for cell formatting.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_format08.xlsx")
local worksheet = workbook:add_worksheet()

local border1 = workbook:add_format({bottom  = 1, bottom_color = "red"})
local border2 = workbook:add_format({top     = 1, top_color    = "red"})
local border3 = workbook:add_format({left    = 1, left_color   = "red"})
local border4 = workbook:add_format({right   = 1, right_color  = "red"})
local border5 = workbook:add_format({border  = 1, border_color = "red"})

worksheet:write_blank("B2",  "", border1)
worksheet:write_blank("B4",  "", border2)
worksheet:write_blank("B6",  "", border3)
worksheet:write_blank("B8",  "", border4)
worksheet:write_blank("B10", "", border5)

workbook:close()
