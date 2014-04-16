----
-- Test cases for xlsxwriter.lua.
--
-- Test the creation of a simple xlsxwriter.lua file with hyperlinks.
-- This example doesn't have any link formatting and tests the relationship
-- linkage code.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook   = Workbook:new("test_hyperlink04.xlsx")
local worksheet1 = workbook:add_worksheet()
local worksheet2 = workbook:add_worksheet()
local worksheet3 = workbook:add_worksheet("Data Sheet")

worksheet1:write_url("A1",  "internal:Sheet2!A1"                                      )
worksheet1:write_url("A3",  "internal:Sheet2!A1:A5"                                   )
worksheet1:write_url("A5",  "internal:'Data Sheet'!D5", nil, "Some text"              )
worksheet1:write_url("E12", "internal:Sheet1!J1"                                      )
worksheet1:write_url("G17", "internal:Sheet2!A1",       nil, "Some text"              )
worksheet1:write_url("A18", "internal:Sheet2!A1",       nil, nil,         "Tool Tip 1")
worksheet1:write_url("A20", "internal:Sheet2!A1",       nil, "More text", "Tool Tip 2")

workbook:close()
