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

local workbook  = Workbook:new("test_hyperlink06.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:write_url("A1", [[external:C:\Temp\foo.xlsx]])
worksheet:write_url("A3", [[external:C:\Temp\foo.xlsx#Sheet1!A1]])
worksheet:write_url("A5", [[external:C:\Temp\foo.xlsx#Sheet1!A1]], nil, "External", "Tip")

workbook:close()
