----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case for cell formatting.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_format10.xlsx")
local worksheet = workbook:add_worksheet()

local border1 = workbook:add_format({bg_color = "red"})
local border2 = workbook:add_format({bg_color = "yellow", pattern = 6})
local border3 = workbook:add_format({bg_color = "yellow",
                                     fg_color = "red",
                                     pattern  = 18})


worksheet:write_blank("B2",  "", border1)
worksheet:write_blank("B4",  "", border2)
worksheet:write_blank("B6",  "", border3)

workbook:close()
