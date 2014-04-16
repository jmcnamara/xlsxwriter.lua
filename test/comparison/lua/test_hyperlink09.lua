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

local workbook  = Workbook:new("test_hyperlink09.xlsx")
local worksheet = workbook:add_worksheet()

worksheet:write_url("A1", [[external:..\foo.xlsx]])
worksheet:write_url("A3", [[external:..\foo.xlsx#Sheet1!A1]])
worksheet:write_url("A5", [[external:\\VBOXSVR\share\foo.xlsx#Sheet1!B2]], nil, [[J:\foo.xlsx#Sheet1!B2]])

workbook:close()

