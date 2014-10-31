----
-- Test cases for xlsxwriter.lua.
--
-- Test the a simple xlsxwriter.lua file with worksheet protection.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_protect02.xlsx")
local worksheet = workbook:add_worksheet()

local unlocked = workbook:add_format({locked = false, hidden = false})
local hidden   = workbook:add_format({locked = false, hidden = true})

worksheet:protect()

worksheet:write("A1", 1)
worksheet:write("A2", 2, unlocked)
worksheet:write("A3", 3, hidden)

workbook:close()

