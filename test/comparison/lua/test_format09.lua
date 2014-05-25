----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case for cell formatting.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_format09.xlsx")
local worksheet = workbook:add_worksheet()

local border1 = workbook:add_format({border    = 7, border_color = "red"})
local border2 = workbook:add_format({diag_type = 1, diag_color   = "red"})
local border3 = workbook:add_format({diag_type = 2, diag_color   = "red"})
local border4 = workbook:add_format({diag_type = 3, diag_color   = "red"})


worksheet:write_blank("B2",  "", border1)
worksheet:write_blank("B4",  "", border2)
worksheet:write_blank("B6",  "", border3)
worksheet:write_blank("B8",  "", border4)

workbook:close()
