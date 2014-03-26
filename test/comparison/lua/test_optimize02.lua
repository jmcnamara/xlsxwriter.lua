----
-- Test cases for xlsxwriter.lua.
--
-- Simple test case to test data writing.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

-- Test writing a workbook in optimization/constant memory mode.
local workbook  = Workbook:new("test_optimize02.xlsx", {constant_memory = true})
local worksheet = workbook:add_worksheet()

worksheet:write("A1", "Hello")
worksheet:write("A2", 123)

-- G1 should be ignored since a later row has already been written.
worksheet:write("G1", "World")

workbook:close()
