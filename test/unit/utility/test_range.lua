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

plan(24)

----
-- Test the utility range function.
--
tests = {
  -- first_row, first_col, last_row, last_col, Range
  {0, 0,   9,       0,     'A1:A10'},
  {1, 2,   8,       2,     'C2:C9'},
  {0, 0,   3,       4,     'A1:E4'},
  {0, 0,   0,       0,     'A1:A1'},
  {0, 0,   0,       1,     'A1:B1'},
  {0, 2,   0,       9,     'C1:J1'},
  {1, 0,   2,       0,     'A2:A3'},
  {9, 0,   1,       24,    'A10:Y2'},
  {7, 25,  9,       26,    'Z8:AA10'},
  {1, 254, 1,       255,   'IU2:IV2'},
  {1, 256, 0,       16383, 'IW2:XFD1'},
  {0, 0,   1048576, 16384, 'A1:XFE1048577'},
}

for _, test in ipairs(tests) do
  got      = utility.range(test[1], test[2], test[3], test[4])
  expected = test[5]
  caption  = string.format(" \trange() -> %s", test[5])
  is(got, expected, caption)
end

----
-- Test the utility range_abs function.
--
tests = {
  -- first_row, first_col, last_row, last_col, Range
  {0, 0,   9,       0,     '$A$1:$A$10'},
  {1, 2,   8,       2,     '$C$2:$C$9'},
  {0, 0,   3,       4,     '$A$1:$E$4'},
  {0, 0,   0,       0,     '$A$1:$A$1'},
  {0, 0,   0,       1,     '$A$1:$B$1'},
  {0, 2,   0,       9,     '$C$1:$J$1'},
  {1, 0,   2,       0,     '$A$2:$A$3'},
  {9, 0,   1,       24,    '$A$10:$Y$2'},
  {7, 25,  9,       26,    '$Z$8:$AA$10'},
  {1, 254, 1,       255,   '$IU$2:$IV$2'},
  {1, 256, 0,       16383, '$IW$2:$XFD$1'},
  {0, 0,   1048576, 16384, '$A$1:$XFE$1048577'},
}

for _, test in ipairs(tests) do
  got      = utility.range_abs(test[1], test[2], test[3], test[4])
  expected = test[5]
  caption  = string.format(" \trange_abs() -> %s", test[5])
  is(got, expected, caption)
end
