----
-- Utility - Utility functions for xlsxwriter.lua.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--
require "xlsxwriter.strict"

local Utility = {}
local char_A = string.byte('A')

----
-- Convert a zero indexed column cell reference to an Excel column string.
--
function Utility.col_to_name_abs(col_num, col_abs)
  local col_str = ''
  local remainder

  col_num = col_num + 1
  col_abs = col_abs and '$' or ''

  while col_num > 0 do
    -- Set remainder from 1 .. 26
    remainder = col_num % 26
    if remainder == 0 then
      remainder = 26
    end

    -- Convert the remainder to a character.
    local col_letter = string.char(char_A + remainder - 1)

    -- Accumulate the column letters, right to left.
    col_str = col_letter .. col_str

    -- Get the next order of magnitude.
    col_num = (col_num - 1) / 26
    col_num = col_num - col_num % 1
  end

  return col_abs .. col_str
end

----
-- Convert a zero indexed row and column cell reference to a A1 style string.
--
function Utility.rowcol_to_cell(row, col)
  row = row + 1
  local col_str = Utility.col_to_name_abs(col, false)
  return col_str .. row
end

----
-- Convert a zero indexed row and column cell reference to a A1 style string
-- with Excel absolute indexing.
--
function Utility.rowcol_to_cell_abs(row, col, row_abs, col_abs)
  row = row + 1
  row_abs = row_abs and '$' or ''
  local col_str = Utility.col_to_name_abs(col, col_abs)
  return col_str .. row_abs .. row
end

----
-- Convert a cell reference in A1 notation to a zero indexed row, column.
--
function Utility.cell_to_rowcol(cell)

  local col_str, row = cell:match("$?(%u+)$?(%d+)")

  -- Convert base26 column string to number.
  local expn = 0
  local col  = 0

  for i = #col_str, 1, -1 do
    local char = col_str:sub(i, i)
    col = col + (string.byte(char) - char_A + 1) * (26 ^ expn)
    expn = expn + 1
  end

  -- Convert 1-index to zero-index
  row = row - 1
  col = col - 1

  return row, col
end

----
-- Convert zero indexed row and col cell refs to a A1:B1 style range string.
--
function Utility.range(first_row, first_col, last_row, last_col)
  local range1 = Utility.rowcol_to_cell(first_row, first_col)
  local range2 = Utility.rowcol_to_cell(last_row,  last_col )
  return range1 .. ':' .. range2
end

----
-- Convert zero indexed row and col cell refs to absolute A1:B1 range string.
--
function Utility.range_abs(first_row, first_col, last_row, last_col)
  local range1 = Utility.rowcol_to_cell_abs(first_row, first_col, true, true)
  local range2 = Utility.rowcol_to_cell_abs(last_row,  last_col,  true, true)
  return range1 .. ':' .. range2
end

return Utility
