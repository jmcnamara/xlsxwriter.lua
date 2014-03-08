----
-- Utility - Utility functions for xlsxwriter.lua.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--
require "xlsxwriter.strict"

local Utility = {}
local char_A = string.byte("A")

local named_colors = {
  ["black"]   = "#000000",
  ["blue"]    = "#0000FF",
  ["brown"]   = "#800000",
  ["cyan"]    = "#00FFFF",
  ["gray"]    = "#808080",
  ["green"]   = "#008000",
  ["lime"]    = "#00FF00",
  ["magenta"] = "#FF00FF",
  ["navy"]    = "#000080",
  ["orange"]  = "#FF6600",
  ["pink"]    = "#FF00FF",
  ["purple"]  = "#800080",
  ["red"]     = "#FF0000",
  ["silver"]  = "#C0C0C0",
  ["white"]   = "#FFFFFF",
  ["yellow"]  = "#FFFF00",
}

----
-- Convert a zero indexed column cell reference to an Excel column string.
--
function Utility.col_to_name_abs(col_num, col_abs)
  local col_str = ""

  col_num = col_num + 1
  col_abs = col_abs and "$" or ""

  while col_num > 0 do
    -- Set remainder from 1 .. 26
    local remainder = col_num % 26
    if remainder == 0 then
      remainder = 26
    end

    -- Convert the remainder to a character.
    local col_letter = string.char(char_A + remainder - 1)

    -- Accumulate the column letters, right to left.
    col_str = col_letter .. col_str

    -- Get the next order of magnitude.
    col_num = math.floor((col_num - 1) / 26)
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
  row_abs = row_abs and "$" or ""
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
  return range1 .. ":" .. range2
end

----
-- Convert zero indexed row and col cell refs to absolute A1:B1 range string.
--
function Utility.range_abs(first_row, first_col, last_row, last_col)
  local range1 = Utility.rowcol_to_cell_abs(first_row, first_col, true, true)
  local range2 = Utility.rowcol_to_cell_abs(last_row,  last_col,  true, true)
  return range1 .. ":" .. range2
end

----
-- Generator for returning table items in sorted order. From PIL 3rd Ed.
--
function Utility.sorted_pairs(sort_table, sort_function)
  local array = {}
  for n in pairs(sort_table) do array[#array + 1] = n end

  table.sort(array, sort_function)

  local i = 0
  return function ()
    i = i + 1
    return array[i], sort_table[array[i]]
  end
end

----
-- Print a non-fatal warning at the highest/calling program stack level.
--
function Utility.warn(...)
  local level = 0
  local info

  -- Find the last highest stack level.
  for i = 1, math.huge do
    info = debug.getinfo(i, "Sl")
    if not info then break end
    level = level + 1
  end

  -- Print warning to stderr at the calling program stack level.
  info = debug.getinfo(level -1, "Sl")
  io.stderr:write(string.format("Warning:\n\t%s:%d: ",
                                info.short_src,
                                info.currentline))
  io.stderr:write(string.format(...))
end

----
-- Convert a Html #RGB or named colour into an Excel ARGB formatted
-- color. Used in conjunction with various xxx_color() methods.
--
function Utility.excel_color(color)
  local rgb = color

  -- Convert named colours.
  if named_colors[color] then rgb = named_colors[color] end

  -- Extract the RBG part of the color.
  rgb = rgb:match("^#(%x%x%x%x%x%x)$")

  if rgb then
    -- Convert the RGB colour to the Excel ARGB format.
    return "FF" .. rgb:upper()
  else
    Utility.warn("Color '%s' is not a valid Excel color.\n", color)
    return "FF000000" -- Return Black as a default on error.
  end

end

return Utility
