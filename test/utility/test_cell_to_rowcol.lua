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

plan(15)

----
-- Test the utility cell_to_rowcol functions.
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
  got      = {utility.cell_to_rowcol(test[3])}
  expected = {test[1], test[2]}
  caption  = string.format(" \tcell_to_rowcol(%s) -> %d, %d",
                           test[3], test[1], test[2])
  eq_array(got, expected, caption)
end
