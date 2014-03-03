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


------------------------------------------------------------------------------
--
-- Public/semi-private methods.
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
    xls_rowmax             = rowmax,
    xls_colmax             = colmax,
    xls_strmax             = strmax,
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
  -- self:_write_cols()

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
    -- carp "Zoom factor scale outside range: 10 <= zoom <= 400"
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
    --carp "Header string must be less than 255 characters"
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
    --carp "Footer string must be less than 255 characters"
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
-- Check that row and col are valid and store max and min values for use in
-- other methods/elements.
--
function Worksheet:_check_dimensions(row, col)

  if row >= xl_rowmax or col >= xl_colmax then
    return false
  end

  -- In optimization mode we don't change dimensions for rows that are
  -- already written.
  if self.optimization then
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
    --self:_write_rows()
    self:_xml_end_tag("sheetData")
  end
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




return Worksheet
