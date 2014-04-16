----
-- Test cases for xlsxwriter.lua.
--
-- Test the creation of a simple xlsxwriter.lua file with strings that
-- require XML escaping.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_escapes01.xlsx")
local worksheet = workbook:add_worksheet("5&4")

worksheet:write_formula("A1", [[=IF(1>2,0,1)]],            nil, 1)
worksheet:write_formula("A2", [[=CONCATENATE("'","<>&")]], nil, [['<>&]])
worksheet:write_formula("A3", [[=1&"b"]],                  nil, [[1b]])
worksheet:write_formula("A4", [[="'"]],                    nil, [[']])
worksheet:write_formula("A5", [[=""""]],                   nil, [["]])
worksheet:write_formula("A6", [[="&" & "&"]],              nil, [[&&]])

worksheet:write_string("A8", [["&<>]])

workbook:close()
