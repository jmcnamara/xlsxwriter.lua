----
-- Tests for the xlsxwriter.lua xml writer class.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"

local utility = require "xlsxwriter.utility"
local expected
local got
local caption
local tests

plan(19)

----
-- Test the utility rowcol_to_cell functions.
--
tests = {
  --row, col, A1 string
  {0, 0, 'A1'},
  {0, 1, 'B1'},
  {0, 2, 'C1'},
  {0, 9, 'J1'},
  {1, 0, 'A2'},
  {2, 0, 'A3'},
  {9, 0, 'A10'},
  {1, 24, 'Y2'},
  {7, 25, 'Z8'},
  {9, 26, 'AA10'},
  {1, 254, 'IU2'},
  {1, 255, 'IV2'},
  {1, 256, 'IW2'},
  {0, 16383, 'XFD1'},
  {1048576, 16384, 'XFE1048577'},
}

for _, test in ipairs(tests) do
  got      = utility.rowcol_to_cell(test[1], test[2])
  expected = test[3]
  caption  = string.format(" \trowcol_to_cell(%d, %d) -> %s",
                           test[1], test[2], test[3])
  is(got, expected, caption)
end


----
-- Test the utility rowcol_to_cell_abs functions.
--
tests = {
   -- row, col, row_abs, col_abs, A1 string
   {0, 0, false, false, 'A1'},
   {0, 0, true,  false, 'A$1'},
   {0, 0, false, true,  '$A1'},
   {0, 0, true,  true,  '$A$1'},
}

for _, test in ipairs(tests) do
   got = utility.rowcol_to_cell_abs(test[1], test[2], test[3], test[4])
   expected = test[5]
   caption = string.format(" \trowcol_to_cell_abs(%d, %d, %s, %s) -> %s",
                           test[1],
                           test[2],
                           tostring(test[3]),
                           tostring(test[4]),
                           test[5])
   is(got, expected, caption)
end
