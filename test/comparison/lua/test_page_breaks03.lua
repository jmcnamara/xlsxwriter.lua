----
-- Test cases for xlsxwriter.lua.
--
-- Test the worksheet set_x_pagebreaks() methods.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_page_breaks03.xlsx")
local worksheet = workbook:add_worksheet()


local breaks = {}

for i = 0, 1025 do
  breaks[i + 1] = i
end

worksheet:set_h_pagebreaks(breaks)

worksheet:write_string("A1", "Foo")

workbook:close()
