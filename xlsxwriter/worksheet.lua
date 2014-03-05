----
-- Worksheet - A class for writing the Excel XLSX Worksheet file.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--
require "xlsxwriter.strict"

local Utility   = require "xlsxwriter.utility"
local Xmlwriter = require "xlsxwriter.xmlwriter"


local xl_rowmax = 1048576
local xl_colmax = 16384
local xl_strmax = 32767


------------------------------------------------------------------------------
--
-- Public methods.
--
------------------------------------------------------------------------------


----
-- The constructor inherits from xmlwriter.lua.
--

local Worksheet = {}
setmetatable(Worksheet,{__index = Xmlwriter})

function Worksheet:new()
  local instance = {
    optimization           = false,
    data_table             = {},

    ext_sheets             = {},
    fileclosed             = false,
    excel_version          = 2007,
    xls_rowmax             = xl_rowmax,
    xls_colmax             = xl_colmax,
    xls_strmax             = xl_strmax,
    dim_rowmin             = nil,
    dim_rowmax             = nil,
    dim_colmin             = nil,
    dim_colmax             = nil,
    colinfo                = {},
    selections             = {},
    hidden                 = false,
    active                 = false,
    tab_color              = 0,
    panes                  = {},
    active_pane            = 3,
    selected               = false,
    page_setup_changed     = false,
    paper_size             = 0,
    orientation            = 1,
    print_options_changed  = false,
    hcenter                = false,
    vcenter                = false,
    print_gridlines        = false,
    screen_gridlines       = true,
    print_headers          = false,
    header_footer_changed  = false,
    header                 = "",
    footer                 = "",
    margin_left            = 0.7,
    margin_right           = 0.7,
    margin_top             = 0.75,
    margin_bottom          = 0.75,
    margin_header          = 0.3,
    margin_footer          = 0.3,
    repeat_rows            = "",
    repeat_cols            = "",
    print_area             = "",
    page_order             = false,
    black_white            = false,
    draft_quality          = false,
    print_comments         = false,
    page_start             = 0,
    fit_page               = false,
    fit_width              = 0,
    fit_height             = 0,
    hbreaks                = {},
    vbreaks                = {},
    protect                = false,
    password               = nil,
    set_cols               = {},
    set_rows               = {},
    zoom                   = 100,
    zoom_scale_normal      = true,
    print_scale            = 100,
    right_to_left          = false,
    show_zeros             = true,
    leading_zeros          = false,
    outline_row_level      = 0,
    outline_col_level      = 0,
    outline_style          = 0,
    outline_below          = true,
    outline_right          = true,
    outline_on             = true,
    outline_changed        = false,
    default_row_height     = 15,
    default_row_zeroed     = false,
    names                  = {},
    write_match            = {},
    merge                  = {},
    has_vml                = false,
    has_comments           = false,
    comments               = {},
    comments_array         = {},
    comments_author        = "",
    comments_visible       = false,
    vml_shape_id           = 1024,
    buttons_array          = {},
    autofilter             = "",
    filter_on              = false,
    filter_range           = {},
    filter_cols            = {},
    col_sizes              = {},
    row_sizes              = {},
    col_formats            = {},
    col_size_changed       = false,
    row_size_changed       = false,
    last_shape_id          = 1,
    rel_count              = 0,
    hlink_count            = 0,
    hlink_refs             = {},
    external_hyper_links   = {},
    external_drawing_links = {},
    external_comment_links = {},
    external_vml_links     = {},
    external_table_links   = {},
    drawing_links          = {},
    charts                 = {},
    images                 = {},
    tables                 = {},
    sparklines             = {},
    shapes                 = {},
    shape_hash             = {},
    has_shapes             = false,
    drawing                = false,
    rstring                = "",
    previous_row           = 0,
  }

  setmetatable(instance, self)
  self.__index = self
  return instance
end


----
-- Assemble and write the XML file.
--
function Worksheet:_assemble_xml_file()

  self:_xml_declaration()

  -- Write the root worksheet element.
  self:_write_worksheet()

  -- Write the worksheet properties.
  -- self:_write_sheet_pr()

  -- Write the worksheet dimensions.
  self:_write_dimension()

  -- Write the sheet view properties.
  self:_write_sheet_views()

  -- Write the sheet format properties.
  self:_write_sheet_format_pr()

  -- Write the sheet column info.
  self:_write_cols()

  -- Write the worksheet data such as rows columns and cells.
  if self.optimization then
    -- self:_write_optimized_sheet_data()
  else
    self:_write_sheet_data()
  end

  -- Write the sheetProtection element.
  -- self:_write_sheet_protection()

  -- Write the autoFilter element.
  -- self:_write_auto_filter()

  -- Write the mergeCells element.
  -- self:_write_merge_cells()

  -- Write the conditional formats.
  -- self:_write_conditional_formats()

  -- Write the dataValidations element.
  -- self:_write_data_validations()

  -- Write the hyperlink element.
  -- self:_write_hyperlinks()

  -- Write the printOptions element.
  -- self:_write_print_options()

  -- Write the worksheet page_margins.
  self:_write_page_margins()

  -- Write the worksheet page setup.
  -- self:_write_page_setup()

  -- Write the headerFooter element.
  -- self:_write_header_footer()

  -- Write the rowBreaks element.
  -- self:_write_row_breaks()

  -- Write the colBreaks element.
  -- self:_write_col_breaks()

  -- Write the drawing element.
  -- self:_write_drawings()

  -- Write the legacyDrawing element.
  -- self:_write_legacy_drawing()

  -- Write the tableParts element.
  -- self:_write_table_parts()

  -- Write the extLst and sparklines.
  -- self:_write_ext_sparklines()

  -- Close the worksheet tag.
  self:_xml_end_tag("worksheet")

  -- Close the XML writer filehandle.
  self:_xml_close()
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
-- Thin wrapper around _write_string() to handle "A1" notation.
--
function Worksheet:write_string(...)
  self:_write_string(self:_convert_cell_args(...))
end

----
-- Thin wrapper around _write_number() to handle "A1" notation.
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

  if not self.data_table[row] then
    self.data_table[row] = {}
  end

  self.data_table[row][col] = {'n',  num, format}

  return 0
end

----
-- Set the worksheet as a selected worksheet, i.e. the worksheet has its tab
-- highlighted.
--
function Worksheet:select()
  -- Selected worksheet can't be hidden.
  self.hidden   = false
  self.selected = true
end

----
-- Set this worksheet as the active worksheet, i.e. the worksheet that is
-- displayed when the workbook is opened. Also set it as selected.
--
function Worksheet:activate()
  -- Active worksheet can't be hidden.
  self.hidden   = false
  self.selected = true
  -- ${ self.activesheet } = self.index
end

----
-- Hide this worksheet.
--
function Worksheet:hide()
  self.hidden = true

  -- A hidden worksheet shouldn't be active or selected.
  self.selected = false
  -- ${ self.activesheet } = false
  -- ${ self.firstsheet }  = false
end

----
-- Set this worksheet as the first visible sheet. This is necessary
-- when there are a large number of worksheets and the activated
-- worksheet is not visible on the screen.
--
function Worksheet:set_first_sheet()
  -- Active worksheet can't be hidden.
  self.hidden = false
  -- ${ self.firstsheet } = self.index
end

----
-- Set the option to hide gridlines on the screen and the printed page.
--
-- This was mainly useful for Excel 5 where printed gridlines were on by
-- default.
--
function Worksheet:hide_gridlines(option)

  -- Default to hiding printed gridlines, option = 1.
  if option == 0 then
    self.print_gridlines       = true
    self.screen_gridlines      = true
    self.print_options_changed = true
  elseif not option or option == 1 then
    self.print_gridlines  = false
    self.screen_gridlines = true
  else
    self.print_gridlines  = false
    self.screen_gridlines = false
  end
end

----
-- Set the worksheet zoom factor.
--
function Worksheet:set_zoom(scale)
  -- Confine the scale to Excel's range
  if scale < 10 or scale > 400 then
    Utility.warn("Zoom factor scale outside range: 10 <= zoom <= 400")
    scale = 100
  end

  self.zoom = math.floor(scale)
end

----
-- Display the worksheet right to left for some eastern versions of Excel.
--
function Worksheet:set_right_to_left()
  self.right_to_left = true
end

----
-- Hide cell zero values.
--
function Worksheet:hide_zero()
  self.show_zeros = false
end

----
-- Set the order in which pages are printed.
--
function Worksheet:print_across(page_order)
  if page_order then
    self.page_order         = true
    self.page_setup_changed = true
  else
    self.page_order = false
  end
end

----
-- Set the start page number.
--
function Worksheet:set_start_page(value)
  self.page_start   = value
  self.custom_start = true
end

----
-- Set the page orientation as portrait.
--
function Worksheet:set_portrait()
  self.orientation        = true
  self.page_setup_changed = true
end

----
-- Set the page orientation as landscape.
--
function Worksheet:set_landscape()
  self.orientation        = false
  self.page_setup_changed = true
end

----
-- Set the page view mode for Mac Excel.
--
function Worksheet:set_page_view()
  self.page_view = true
end

----
-- Set the colour of the worksheet tab.
--
function Worksheet:set_tab_color(color)
  self.tab_color = color
end

----
-- Set the paper type. Ex. 1 = US Letter, 9 = A4
--
function Worksheet:set_paper(paper_size)
  self.paper_size         = paper_size
  self.page_setup_changed = 1
end

----
-- Set the page header caption and optional margin.
--
function Worksheet:set_header(header, margin)

  if #header >= 255 then
    Utility.warn("Header string must be less than 255 characters")
    return
  end

  self.header                = header
  self.margin_header         = margin and margin or 0.3
  self.header_footer_changed = 1
end

----
-- Set the page footer caption and optional margin.
--
function Worksheet:set_footer(footer, margin)

  if #footer >= 255 then
    Utility.warn("Footer string must be less than 255 characters")
    return
  end

  self.footer                = footer
  self.margin_footer         = margin and margin or 0.3
  self.header_footer_changed = 1
end

----
-- Set the page margins in inches.
--
function Worksheet:set_margins(left, right, top, bottom)
  self.margin_left   = left   and left   or 0.7
  self.margin_right  = right  and right  or 0.7
  self.margin_top    = top    and top    or 0.75
  self.margin_bottom = bottom and bottom or 0.75
end

----
-- Thin wrapper around _set_column() to handle "A:Z" notation.
--
function Worksheet:set_column(...)
  self:_set_column(self:_convert_column_args(...))
end

----
-- Set the width of a single column or a range of columns.
--
function Worksheet:_set_column(firstcol, lastcol, width, format, options)

  -- Ensure 2nd col is larger than first. Also for KB918419 bug.
  if firstcol > lastcol then
    firstcol, lastcol = lastcol, firstcol
  end

  -- Set the optional column values.
  options = options or {}
  local hidden    = options["hidden"]
  local collapsed = options["collapsed"]
  local level     = options["level"] or 0

  -- Check that cols are valid and store max and min values with default row.
  -- NOTE: The check shouldn't modify the row dimensions and should only modify
  --       the column dimensions in certain cases.
  local ignore_row = true
  local ignore_col = true

  if format or (width and hidden) then
    ignore_col = false
  end


  if not self:_check_dimensions(0, firstcol, ignore_row, ignore_col) then
    return -1
  end

  if not self:_check_dimensions(0, lastcol, ignore_row, ignore_col) then
    return -1
  end


  -- Set the limits for the outline levels (0 <= x <= 7).
  if level < 0 then level = 0  end
  if level > 7 then level = 7  end

  if level > self.outline_col_level then
    self.outline_col_level = level
  end


  -- Store the column data based on the first column. Padded for sorting.
  self.colinfo[string.format("%05d", firstcol)] = {["firstcol"]  = firstcol,
                                                   ["lastcol"]   = lastcol,
                                                   ["width"]     = width,
                                                   ["format"]    = format,
                                                   ["hidden"]    = hidden,
                                                   ["level"]     = level,
                                                   ["collapsed"] = collapsed}

  -- Store the column change to allow optimisations.
  self.col_size_changed = 1

  -- Store the col sizes for use when calculating image vertices taking
  -- hidden columns into account. Also store the column formats.
  if hidden then width = 0 end

  for col = firstcol, lastcol do
    self.col_sizes[col] = width
    if format then
      self.col_formats[col] = format
    end
  end
end



------------------------------------------------------------------------------
--
-- Internal methods.
--
------------------------------------------------------------------------------

----
-- Decorator function to convert "A1" notation in cell method calls
-- to the default row/col notation.
--
function Worksheet:_convert_cell_args(...)
  if type(...) == "string" then
    -- Convert "A1" style cell to row, col.
    local cell = ...
    local row, col = Utility.cell_to_rowcol(cell)
    return row, col, unpack({...}, 2)
  else
    -- Parameters are already in row, col format.
    return ...
  end
end

----
-- Decorator function to convert "A:Z" column range calls to column numbers.
--
function Worksheet:_convert_column_args(...)
  if type(...) == "string" then
    -- Convert "A:Z" style range to col, col.
    local range = ...
    local range_start, range_end = range:match("(%S+):(%S+)")
    local _, col_1 = Utility.cell_to_rowcol(range_start .. "1")
    local _, col_2 = Utility.cell_to_rowcol(range_end   .. "1")
    return col_1, col_2, unpack({...}, 2)
  else
    -- Parameters are already in col, col format.
    return ...
  end
end




----
-- Check that row and col are valid and store max and min values for use in
-- other methods/elements.
--
-- The ignore_row/ignore_col flags is used to indicate that we wish to
-- perform the dimension check without storing the value.
--
-- The ignore flags are use by set_row() and data_validate.
--
function Worksheet:_check_dimensions(row, col, ignore_row, ignore_col)

  if row >= xl_rowmax or col >= xl_colmax then
    return false
  end

  -- In optimization mode we don't change dimensions for rows that are
  -- already written.
  if self.optimization and not ignore_row and not ignore_col then
    if row < self.previous_row then
      return -2
    end
  end

  if not ignore_row then
    if not self.dim_rowmin or row < self.dim_rowmin then
      self.dim_rowmin = row
    end

    if not self.dim_rowmax or row > self.dim_rowmax then
      self.dim_rowmax = row
    end
  end

  if not ignore_col then
    if not self.dim_colmin or col < self.dim_colmin then
      self.dim_colmin = col
    end

    if not self.dim_colmax or col > self.dim_colmax then
      self.dim_colmax = col
    end
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

  local schema   = "http://schemas.openxmlformats.org/"
  local xmlns    = schema .. "spreadsheetml/2006/main"
  local xmlns_r  = schema .. "officeDocument/2006/relationships"

  local attributes = {
    {["xmlns"]   = xmlns},
    {["xmlns:r"] = xmlns_r},
  }

  self:_xml_start_tag("worksheet", attributes)
end

----
-- Write the <dimension> element. This specifies the range of cells in the
-- Worksheet. As a special case, empty spreadsheets use "A1" as a range.
--
function Worksheet:_write_dimension()
  local ref = ""

  if not self.dim_rowmin and not self.dim_colmin then
    -- If the min dims are undefined then no dimensions have been set
    -- and we use the default "A1".
    ref = "A1"
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

  self:_xml_empty_tag("dimension", {{["ref"] = ref}})
end


----
-- Write the <sheetViews> element.
--
function Worksheet:_write_sheet_views()


  local attributes = {}

  self:_xml_start_tag("sheetViews", attributes)
  self:_write_sheet_view()
  self:_xml_end_tag("sheetViews")
end

----
-- Write the <sheetView> element.
--
-- Sample structure:
--     <sheetView
--         showGridLines="0"
--         showRowColHeaders="0"
--         showZeros="0"
--         rightToLeft="1"
--         tabSelected="1"
--         showRuler="0"
--         showOutlineSymbols="0"
--         view="pageLayout"
--         zoomScale="121"
--         zoomScaleNormal="121"
--         workbookViewId="0"
--      />
--
function Worksheet:_write_sheet_view()

  local gridlines        = self.screen_gridlines
  local show_zeros       = self.show_zeros
  local right_to_left    = self.right_to_left
  local tab_selected     = self.selected
  local view             = self.page_view
  local zoom             = self.zoom
  local workbook_view_id = 0
  local attributes       = {}

  -- Hide screen gridlines if required
  if not gridlines then
    table.insert(attributes, {["showGridLines"] = "0"})
  end

  -- Hide zeroes in cells.
  if not show_zeros then
    table.insert(attributes, {["showZeros"] = "0"})
  end

  -- Display worksheet right to left for Hebrew, Arabic and others.
  if right_to_left then
    table.insert(attributes, {["rightToLeft"] = "1"})
  end

  -- Show that the sheet tab is selected.
  if tab_selected then
    table.insert(attributes, {["tabSelected"] = "1"})
  end

  -- Turn outlines off. Also required in the outlinePr element.
  if not self.outline_on then
    table.insert(attributes, {["showOutlineSymbols"] = "0"})
  end

  -- Set the page view/layout mode if required.
  if view then
    table.insert(attributes, {["view"] = "pageLayout"})
  end

  -- Set the zoom level.
  if zoom ~= 100 then
    if not view then
      table.insert(attributes, {["zoomScale"] = zoom})
    end
    if self.zoom_scale_normal then
      table.insert(attributes, {["zoomScaleNormal"] = zoom})
    end
  end

  table.insert(attributes, {["workbookViewId"] = workbook_view_id})

  self:_xml_empty_tag("sheetView", attributes)
end

----
-- Write the <pageMargins> element.
--
function Worksheet:_write_page_margins()
  local attributes = {
    {["left"]   = self.margin_left},
    {["right"]  = self.margin_right},
    {["top"]    = self.margin_top},
    {["bottom"] = self.margin_bottom},
    {["header"] = self.margin_header},
    {["footer"] = self.margin_footer},
  }

  self:_xml_empty_tag("pageMargins", attributes)
end

----
-- Write the <sheetFormatPr> element.
--
function Worksheet:_write_sheet_format_pr()

  local base_col_width     = 10
  local default_row_height = self.default_row_height
  local row_level          = self.outline_row_level
  local col_level          = self.outline_col_level
  local zero_height        = self.default_row_zeroed

  local attributes = {{["defaultRowHeight"] = default_row_height}}

  if self.default_row_height ~= 15 then
    table.insert(attributes, {["customHeight"] = "1"})
  end

  if self.default_row_zeroed then
    table.insert(attributes, {["zeroHeight"] = "1"})
  end

  if row_level > 0 then
    table.insert(attributes, {["outlineLevelRow"] = row_level})
  end

  if col_level > 0 then
    table.insert(attributes, {["outlineLevelCol"] = col_level})
  end

  if self.excel_version == 2010 then
    table.insert(attributes, {["x14ac:dyDescent"] = "0.25"})
  end

  self:_xml_empty_tag("sheetFormatPr", attributes)
end

----
-- Write the <sheetData> element.
--
function Worksheet:_write_sheet_data()
  if not self.dim_rowmin then
    -- If the dimensions aren't defined then there is no data to write.
    self:_xml_empty_tag("sheetData")
  else
    self:_xml_start_tag("sheetData")
    self:_write_rows()
    self:_xml_end_tag("sheetData")
  end
end



----
-- Write out the worksheet data as a series of rows and cells.
--
function Worksheet:_write_rows()
  -- Calculate the row span attributes.
  self:_calculate_spans()

  for row_num = self.dim_rowmin, self.dim_rowmax do

    -- Only write rows if they contain row formatting, cell data or a comment.
    if self.set_rows[row_num] or self.data_table[row_num] or self.comments[row_num] then

      local span_index = math.floor(row_num / 16)
      local span       = self.row_spans[span_index]

      -- Write the cells if the row contains data.
      if self.data_table[row_num] then

        if not self.set_rows[row_num] then
          self:_write_row(row_num, span)
        else
          self:_write_row(row_num, span, unpack(self.set_rows[row_num]))
        end

        for col_num = self.dim_colmin, self.dim_colmax do
          if self.data_table[row_num][col_num] then
            self:_write_cell(row_num, col_num, self.data_table[row_num][col_num])
          end
        end

        self:_xml_end_tag("row")

      elseif self.comments[row_num] then
        self:_write_empty_row(row_num, span, unpack(self.set_rows[row_num]))
      else
        -- Row attributes only.
        self:_write_empty_row(row_num, span, unpack(self.set_rows[row_num]))
      end
    end
  end
end

----
-- Write out the worksheet data as a single row with cells. This method is
-- used when memory optimisation is on. A single row is written and the data
-- table is reset. That way only one row of data is kept in memory at any one
-- time. We don't write span data in the optimised case since it is optional.
--
function Worksheet:_write_single_row(current_row)

  local row_num = self.previous_row

  -- Set the new previous row as the current row.
  self.previous_row = current_row

  -- Only write rows if they contain row formatting, cell data or a comment.
  if self.set_rows[row_num] or self.data_table[row_num] or self.comments[row_num] then

    -- Write the cells if the row contains data.
    if self.data_table[row_num] then

      if not self.set_rows[row_num] then
        self:_write_row(row_num)
      else
        self:_write_row(row_num, nil, unpack(self.set_rows[row_num]))
      end

      for col_num = self.dim_colmin, self.dim_colmax do
        if self.data_table[row_num][col_num] then
          self:_write_cell(row_num, col_num, self.data_table[row_num][col_num])
        end
      end

      self:_xml_end_tag("row")
    else
      -- Row attributes or comments only.
      self:_write_empty_row(row_num, nil, unpack(self.set_rows[row_num]))
    end

    -- Reset table.
    self.data_table = {}
  end
end


----
-- Calculate the "spans" attribute of the <row> tag. This is an XLSX
-- optimisation and isn't strictly required. However, it makes comparing
-- files easier.
--
-- The span is the same for each block of 16 rows.
--
function Worksheet:_calculate_spans()
  local spans = {}
  local span_min
  local span_max

  for row_num = self.dim_rowmin, self.dim_rowmax do
    -- Calculate spans for cell data.
    if self.data_table[row_num] then
      for col_num = self.dim_colmin, self.dim_colmax do
        if self.data_table[row_num][col_num] then
          if not span_min then
            span_min = col_num
            span_max = col_num
          else
            if col_num < span_min then
              span_min = col_num
            end
            if col_num > span_max then
              span_max = col_num
            end
          end
        end
      end
    end

    -- Calculate spans for comments.
    if self.comments[row_num] then
      for col_num = self.dim_colmin, self.dim_colmax do
        if self.comments[row_num][col_num] then
          if not span_min then
            span_min = col_num
            span_max = col_num
          else
            if col_num < span_min then
              span_min = col_num
            end
            if col_num > span_max then
              span_max = col_num
            end
          end
        end
      end
    end

    if (row_num + 1) % 16 == 0  or row_num == self.dim_rowmax then
      local span_index = math.floor(row_num / 16)

      if span_min then
        span_min = span_min + 1
        span_max = span_max + 1
        spans[span_index] = string.format("%d:%d", span_min, span_max)

        span_min = nil
        span_max = nil
      end
    end
  end

  self.row_spans = spans
end

----
-- Write the <row> element.
--
function Worksheet:_write_row(r, spans, height, format, hidden, level, collapsed, empty_row)

  local xf_index = 0

  if not height then
    height = self.default_row_height
  end

  local attributes = {{["r"] = r + 1}}

  -- Get the format index.
  if format then
    -- xf_index = format:get_xf_index()
    xf_index = format
  end

  -- TODO. Rewrite all as [#attributes + 1].
  if spans then
    table.insert(attributes, {["spans"] = spans})
  end

  if xf_index > 0 then
    table.insert(attributes, {["s"] = xf_index})
  end

  if format then
    table.insert(attributes, {["customFormat"] = "1"})
  end

  if height ~= 15 then
    table.insert(attributes, {["ht"] = height})
  end

  if hidden then
    table.insert(attributes, {["hidden"] = "1"})
  end

  if height ~= 15 then
    table.insert(attributes, {["customHeight"] = 1})
  end

  if level then
    table.insert(attributes, {["outlineLevel"] = level})
  end

  if collapsed then
    table.insert(attributes, {["collapsed"]    = "1"})
  end

  if self.excel_version == 2010 then
    table.insert(attributes, {["x14ac:dyDescent"] = "0.25"})
  end

  if empty_row then
    self:_xml_empty_tag_unencoded("row", attributes)
  else
    self:_xml_start_tag_unencoded("row", attributes)
  end
end

----
-- Write and empty <row> element, i.e., attributes only, no cell data.
--
function Worksheet:_write_empty_row(r, spans, height, format, hidden, level, collapsed)
  -- Set the $empty_row parameter.
  local empty_row = 1
  self:_write_row(r, spans, height, format, hidden, level, collapsed, empty_row)
end


----
-- Write the <cell> element. This is the innermost loop so efficiency is
-- important where possible. The basic methodology is that the data of every
-- cell type is passed in as follows:
--
--      (row, col, cell)
--
-- The cell is a table containing the following structure in all types:
--
--     {cell_type, token, xf, args}
--
-- Where cell_type: represents the cell type, such as string, number, formula.
--       token:     is the actual data for the string, number, formula, etc.
--       xf:        is the XF format object.
--       args:      additional args relevant to the specific data type.
--
function Worksheet:_write_cell(row, col, cell)

  local cell_type = cell[1]
  local token     = cell[2]
  local xf        = cell[3]
  local xf_index  = 0

  -- Get the format index.
  if xf then
    xf_index = xf
  end

  local range = Utility.rowcol_to_cell(row, col)
  local attributes = {{["r"] = range}}
  -- TODO. Rewrite all as [#attributes + 1].

  -- Add the cell format index.
  if xf_index > 0 then

    table.insert(attributes, {["s"] = xf_index})

  elseif self.set_rows[row] and self.set_rows[row][1] then

    local row_xf = self.set_rows[row][1]
    table.insert(attributes, {["s"] = row_xf})

  elseif self.col_formats[col] then

    local col_xf = self.col_formats[col]
    table.insert(attributes, {["s"] = col_xf})
  end

  -- Write the various cell types.
  if cell_type == "n" then
    -- Write a number.
    self:_xml_number_element(token, attributes)

  elseif cell_type == "s" then
    -- Write a string.
    if not self.optimization then
      self:_xml_string_element(token, attributes)
    else
      local str = token
      -- Escape control characters. See SharedString.pm for details.
      --str =~ s/(_x[0-9a-fA-F]{4}_)/_x005F1/g
      --str =~ s/([\x00-\x08\x0B-\x1F])/sprintf "_x04X_", ord(1)/eg

      -- Write any rich strings without further tags.
      -- if str =~ m{^<r>} and str =~ m{</r>$} then
      --   self:_xml_rich_inline_string(str, attributes)
      -- else

      -- Add attribute to preserve leading or trailing whitespace.
      local preserve = false
      if string.match(str, "^%s") or string.match(str, "%s$") then
        preserve = true
      end
      self:_xml_inline_string(str, preserve, attributes)
    end

  elseif cell_type == "f" then

    -- Write a formula.
    local value = cell[4] or 0

    -- Check if the formula value is a string.
    if type(value) == "string" then
      table.insert(attributes, {["t"] = "str"})
      value = Utility._escape_data(value)
    end

    self:_xml_formula_element(token, value, attributes)

  elseif cell_type == "a" then

    -- Write an array formula.
    self:_xml_start_tag("c", attributes)
    self:_write_cell_array_formula(token, cell[4])
    self:_write_cell_value(cell[5])
    self:_xml_end_tag("c")

  elseif cell_type == "b" then
    -- Write a empty cell.
    self:_xml_empty_tag("c", attributes)
  end
end

----
-- Write the cell value <v> element.
--
function Worksheet:_write_cell_value(value)
  self:_xml_data_element("v", value or '')
end

----
-- Write the cell formula <f> element.
--
function Worksheet:_write_cell_formula(formula)
  self:_xml_data_element("f", formula or '')
end

----
-- Write the cell array formula <f> element.
--
function Worksheet:_write_cell_array_formula(formula, range)
  local attributes = {{["t"] = "array"}, {["ref"] = range}}
  self:_xml_data_element("f", formula, attributes)
end

----
-- Write the <cols> element and <col> sub elements.
--
function Worksheet:_write_cols()

  -- Return unless some column have been formatted.
  if next(self.colinfo) == nil then return end

  self:_xml_start_tag("cols")

  for row, colinfo in Utility.sorted_pairs(self.colinfo) do
    self:_write_col_info(self.colinfo[row])
  end

  self:_xml_end_tag("cols")
end

----
-- Write the <col> element.
--
function Worksheet:_write_col_info(colinfo)

  local firstcol     = colinfo["firstcol"]
  local lastcol      = colinfo["lastcol"]
  local width        = colinfo["width"]
  local format       = colinfo["format"]
  local hidden       = colinfo["hidden"]
  local level        = colinfo["level"]
  local collapsed    = colinfo["collapsed"]
  local custom_width = true
  local xf_index     = 0

  -- Get the format index.
  if format and format > 0 then
    xf_index = format
  end

  -- Set the Excel default col width.
  if not width then
    if not hidden then
      width        = 8.43
      custom_width = false
    else
      width = 0
    end
  else

    -- Width is defined but same as default.
    if width == 8.43 then
      custom_width = false
    end
  end

  -- Convert column width from user units to character width.
  local max_digit_width = 7    -- For Calabri 11.
  local padding         = 5

  if width > 0 then
    if width < 1 then
      width = math.floor((math.floor(width*(max_digit_width + padding) + 0.5))
                           / max_digit_width * 256) / 256
    else
      width = math.floor((math.floor(width*max_digit_width + 0.5) + padding)
                           / max_digit_width * 256) / 256
    end
  end

  local attributes = {
    {["min"]   = firstcol + 1},
    {["max"]   = lastcol  + 1},
    {["width"] = width},
  }

  if xf_index > 0 then
    table.insert(attributes, {["style"] = xf_index})
  end

  if hidden then
    table.insert(attributes, {["hidden"] = "1"})
  end

  if custom_width then
    table.insert(attributes, {["customWidth"] = "1"})
  end

  if level > 0 then
    table.insert(attributes, {["outlineLevel"] = level})
  end

  if collapsed then
    table.insert(attributes, {["collapsed"] = "1"})
  end

  self:_xml_empty_tag("col", attributes)
end


return Worksheet
