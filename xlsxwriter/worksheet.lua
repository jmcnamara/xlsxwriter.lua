----
-- Worksheet - A class for writing the Excel XLSX Worksheet file.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Utility   = require "xlsxwriter.utility"
local Xmlwriter = require "xlsxwriter.xmlwriter"


local xl_rowmax = 1048576
local xl_colmax = 16384
local xl_strmax = 32767


----
-- The constructor inherits from xmlwriter.lua.
--
local Worksheet = Xmlwriter:new{
};



----
-- Decorator function to convert 'A1' notation in cell method calls
-- to the default row/col notation.
--
function Worksheet:_convert_cell_args(...)
  if type(...) == "string" then
    -- Convert 'A1' style cell to row, col.
    local cell = ...
    local row, col = Utility.cell_to_rowcol(cell)
    return row, col, unpack({...}, 2)
  else
    -- Parameters are already in row, col format.
    return ...
  end
end

----
-- Write data to a Worksheet cell by calling the appropriate _write_*()
-- method based on the type of data being passed.
--
function Worksheet:write(...)
  local row, col, token, format = self:_convert_cell_args(...)

  if type(token) == "string" then
    self:_write_string(row, col, token, format)
  else
    self:_write_number(row, col, token, format)
  end
end

----
-- Thin wrapper around _write_string() to handle 'A1' notation.
--
function Worksheet:write_string(...)
  self:_write_string(self:_convert_cell_args(...))
end

----
-- Thin wrapper around _write_number() to handle 'A1' notation.
--
function Worksheet:write_number(...)
  self:_write_string(self:_convert_cell_args(...))
end


----
-- Write a string to a Worksheet cell.
--
function Worksheet:_write_string(row, col, str, format)

  if not self:_check_dimensions(row, col) then
    return -1
  end

end

----
-- Write a number to a Worksheet cell.
--
function Worksheet:_write_number(row, col, num, format)

  if not self:_check_dimensions(row, col) then
    return -1
  end

  self:_check_dimensions(row, col)
end


----
-- Check that row and col are valid and store max and min values for use in
-- other methods/elements.
--
function Worksheet:_check_dimensions(row, col)

  if row >= xl_rowmax or col >= xl_colmax then
    return false
  end

  -- In optimization mode we don't change dimensions for rows that are
  -- already written.
  if self.optimization == 1 then
    if row < self.previous_row then
      return false
    end
  end

  if not self.dim_rowmin or row < self.dim_rowmin then
    self.dim_rowmin = row
  end

  if not self.dim_rowmax or row > self.dim_rowmax then
    self.dim_rowmax = row
  end

  if not self.dim_colmin or col < self.dim_colmin then
    self.dim_colmin = col
  end

  if not self.dim_colmax or col > self.dim_colmax then
    self.dim_colmax = col
  end

  return true
end




------------------------------------------------------------------------------
--
-- XML writing methods.
--
------------------------------------------------------------------------------

----
-- Write the <Worksheet> element. This is the root element of Worksheet.
--
function Worksheet:_write_worksheet()

  local schema   = 'http://schemas.openxmlformats.org/'
  local xmlns    = schema .. 'spreadsheetml/2006/main'
  local xmlns_r  = schema .. 'officeDocument/2006/relationships'

  local attributes = {
    {['xmlns']   = xmlns},
    {['xmlns:r'] = xmlns_r},
  }

  self:_xml_start_tag('worksheet', attributes)
end

----
-- Write the <dimension> element. This specifies the range of cells in the
-- Worksheet. As a special case, empty spreadsheets use 'A1' as a range.
--
function Worksheet:_write_dimension()
  local ref = ''

  if not self.dim_rowmin and not self.dim_colmin then
    -- If the min dims are undefined then no dimensions have been set
    -- and we use the default 'A1'.
    ref = 'A1'
  elseif not self.dim_rowmin and self.dim_colmin then
    -- If the row dims aren't set but the column dims are then they
    -- have been changed via set_column().
    if self.dim_colmin == self.dim_colmax then
      -- The dimensions are a single cell and not a range.
      ref = Utility.rowcol_to_cell(0, self.dim_colmin)
    else
      -- The dimensions are a cell range.
      ref  = Utility.range(0, self.dim_colmin, 0, self.dim_colmax)
    end
  elseif self.dim_rowmin == self.dim_rowmax
     and self.dim_colmin == self.dim_colmax then
    -- The dimensions are a single cell and not a range.
    ref = Utility.rowcol_to_cell(self.dim_rowmin, self.dim_colmin)
  else
    -- The dimensions are a cell range.
    ref = Utility.range(self.dim_rowmin, self.dim_colmin,
                        self.dim_rowmax, self.dim_colmax)
  end

  self:_xml_empty_tag('dimension', {{['ref'] = ref}})
end


return Worksheet
