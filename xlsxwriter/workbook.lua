----
-- Workbook - A class for writing the Excel XLSX Workbook file.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--
require "xlsxwriter.strict"

local Xmlwriter = require "xlsxwriter.xmlwriter"
local Worksheet = require "xlsxwriter.worksheet"
local Format    = require "xlsxwriter.format"

------------------------------------------------------------------------------
--
-- Constructor.
--
------------------------------------------------------------------------------

-- The constructor inherits from xmlwriter.lua.
local Workbook = {}
setmetatable(Workbook,{__index = Xmlwriter})

function Workbook:new()
  local instance = {

    worksheet_meta     = {activesheet = 0, firstsheet = 0},
    selected           = 0,
    fileclosed         = false,
    filehandle         = false,
    internal_fh        = false,
    sheet_name         = 'Sheet',
    chart_name         = 'Chart',
    worksheet_count    = 0,
    sheetname_count    = 0,
    chartname_count    = 0,
    worksheets         = {},
    charts             = {},
    drawings           = {},
    sheetnames         = {},
    formats            = {},
    xf_formats         = {},
    xf_format_indices  = {},
    dxf_formats        = {},
    dxf_format_indices = {},
    palette            = {},
    font_count         = 0,
    num_format_count   = 0,
    defined_names      = {},
    named_ranges       = {},
    custom_colors      = {},
    doc_properties     = {},
    localtime          = 0,
    num_vml_files      = 0,
    num_comment_files  = 0,
    x_window           = 240,
    y_window           = 15,
    window_width       = 16095,
    window_height      = 9660,
    tab_ratio          = 500,
    str_table          = {},
    vba_project        = false,
    vba_codename       = false,
    image_types        = {},
    images             = {},
    border_count       = 0,
    fill_count         = 0,
    drawing_count      = 0,
    calc_mode          = "auto",
    calc_on_load       = true,
  }

  setmetatable(instance, self)
  self.__index = self


  return instance
end

------------------------------------------------------------------------------
--
-- Public methods.
--
------------------------------------------------------------------------------

----
-- Assemble and write the XML file.
--
function Workbook:_assemble_xml_file()

  -- Prepare format object for passing to Style.pm.
  --self:_prepare_format_properties()

  self:_xml_declaration()

  -- Write the root workbook element.
  self:_write_workbook()

  -- Write the XLSX file version.
  self:_write_file_version()

  -- Write the workbook properties.
  self:_write_workbook_pr()

  -- Write the workbook view properties.
  self:_write_book_views()

  -- Write the worksheet names and ids.
  self:_write_sheets()

  -- Write the workbook defined names.
  self:_write_defined_names()

  -- Write the workbook calculation properties.
  self:_write_calc_pr()

  -- Close the workbook tag.
  self:_xml_end_tag("workbook")

  -- Close the XML writer filehandle.
  self:_xml_close()

end

----
-- Add a new worksheet to the Excel workbook.
--
-- Args:
--     name: The worksheet name. Defaults to 'Sheet1', etc.
--
-- Returns:
--     Reference to a worksheet object.
--
function Workbook:add_worksheet(name)
  return self:_add_sheet(name)
end








----
-- Calls finalization methods.
--
function Workbook:close()
  if not self.fileclosed then
    self.fileclosed = true
    self:_store_workbook()
  end
end


------------------------------------------------------------------------------
--
-- Internal methods.
--
------------------------------------------------------------------------------

----
-- Utility for shared code in add_worksheet() and add_chartsheet().
--
function Workbook:_add_sheet(name, is_chartsheet)

  local sheet_index = self.worksheet_count
  local name        = self:_check_sheetname(name)

  -- Initialisation data to pass to the worksheet.
  local init_data = {
    ["name"]                = name,
    ["index"]               = sheet_index,
    ["str_table"]           = self.str_table,
    ["worksheet_meta"]      = self.worksheet_meta,
    ["optimization"]        = self.optimization,
    ["tmpdir"]              = self.tmpdir,
    ["date_1904"]           = self.date_1904,
    ["strings_to_numbers"]  = self.strings_to_numbers,
    ["strings_to_formulas"] = self.strings_to_formulas,
    ["strings_to_urls"]     = self.strings_to_urls,
    ["default_date_format"] = self.default_date_format,
    ["default_url_format"]  = self.default_url_format,
  }

  local worksheet
  if is_chartsheet then
    -- worksheet = Chartsheet:new()
  else
    worksheet = Worksheet:new()
  end

  worksheet:_initialize(init_data)

  self.worksheet_count = self.worksheet_count + 1
  self.worksheets[self.worksheet_count] = worksheet
  self.sheetnames[self.worksheet_count] = name

  return worksheet
end

----
-- Assemble worksheets into a workbook.
--
function Workbook:_store_workbook()

  -- Add a default worksheet if non have been added.
  if #self.worksheets == 0 then
    self:add_worksheet()
  end

  -- Ensure that at least one worksheet has been selected.
  if self.activesheet == 0 then
    self.worksheets[1].selected = 1
    self.worksheets[1].hidden   = 0
  end

  -- Set the active sheet.
  for _, sheet in ipairs(self.worksheets) do
    if sheet.index == self.activesheet then
      sheet.active = 1
    end
  end

  -- Prepare the worksheet VML elements such as comments and buttons.
  --self:_prepare_vml_objects()

  -- Set the defined names for the worksheets such as Print Titles.
  --self:_prepare_defined_names()

  -- Prepare the drawings, charts and images.
  --self:_prepare_drawings()

  -- Add cached data to charts.
  --self:_add_chart_data()

  -- Prepare the worksheet tables.
  --self:_prepare_tables()

  -- Package the workbook.
  -- TODO

end


----
-- Check for valid worksheet names. We check the length, if it contains any
-- invalid characters and if the name is unique in the workbook.
--
function Workbook:_check_sheetname(name, is_chartsheet)

  -- Increment the Sheet/Chart number used for default sheet names below.
  if is_chartsheet then
    self.chartname_count = self.chartname_count + 1
  else
    self.sheetname_count = self.sheetname_count + 1
  end

  -- Supply default Sheet/Chart name if none has been defined.
  if not name or name == "" then
    if is_chartsheet then
      name = self.chart_name .. tostring(self.chartname_count)
    else
      name = self.sheet_name .. tostring(self.sheetname_count)
    end
  end

  -- Check that sheet name is <= 31. Excel limit.
  -- TODO. Need to add a UTF-8 length check.
  assert(#name <= 31, string.format("Sheetname '%s' must be <= 31 chars", name))

  -- Check that sheetname doesn't contain any invalid characters
  if name:match('[%[%]:%*%?/\\]') then
    error("Invalid character '[]:*?/\\' in worksheet name: " .. name)
  end

  -- Check that the worksheet name doesn't already exist since this is a fatal
  -- error in Excel 97+. The check must also exclude case insensitive matches.
  for _, worksheet in ipairs(self.worksheets) do
    local name_a = name
    local name_b = worksheet.name

    if name_a:lower() == name_b:lower() then
      error(string.format(
              "Worksheet name "%s", with case ignored, is already used.",
              name))
    end
  end

  return name
end




------------------------------------------------------------------------------
--
-- XML writing methods.
--
------------------------------------------------------------------------------

----
-- Write <workbook> element.
--
function Workbook:_write_workbook()
  local schema  = "http://schemas.openxmlformats.org"
  local xmlns   = schema .. "/spreadsheetml/2006/main"
  local xmlns_r = schema .. "/officeDocument/2006/relationships"

  local attributes = {
    {["xmlns"]   = xmlns},
    {["xmlns:r"] = xmlns_r},
  }

  self:_xml_start_tag("workbook", attributes)
end

----
-- Write the <fileVersion> element.
--
function Workbook:_write_file_version()
  local app_name      = "xl"
  local last_edited   = "4"
  local lowest_edited = "4"
  local rup_build     = "4505"

  local attributes = {
    {["appName"]      = app_name},
    {["lastEdited"]   = last_edited},
    {["lowestEdited"] = lowest_edited},
    {["rupBuild"]     = rup_build},
  }

  if self.vba_project then
    table.insert(attributes,
                 {["codeName"] = "{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}"})
  end

  self:_xml_empty_tag("fileVersion", attributes)
end

----
-- Write <workbookPr> element.
--
function Workbook:_write_workbook_pr()
  local date_1904             = self.date_1904
  local default_theme_version = 124226
  local codename              = self.vba_codename
  local attributes ={}

  if codename then
     table.insert(attributes, {["codeName"] = codename})
  end

  if date_1904 then
     table.insert(attributes, {["date1904"] = 1})
  end

  table.insert(attributes, {["defaultThemeVersion"] = default_theme_version})

  self:_xml_empty_tag("workbookPr", attributes)
end

----
-- Write <bookViews> element.
--
function Workbook:_write_book_views()
  self:_xml_start_tag("bookViews")
  self:_write_workbook_view()
  self:_xml_end_tag("bookViews")
end

----
-- Write <workbookView> element.
--
function Workbook:_write_workbook_view()
  local x_window      = self.x_window
  local y_window      = self.y_window
  local window_width  = self.window_width
  local window_height = self.window_height
  local tab_ratio     = self.tab_ratio
  local active_tab    = self.worksheet_meta.activesheet
  local first_sheet   = self.worksheet_meta.firstsheet

  local attributes = {
    {["xWindow"]      = x_window},
    {["yWindow"]      = y_window},
    {["windowWidth"]  = window_width},
    {["windowHeight"] = window_height},
  }

  -- Store the tabRatio attribute when it isn't the default.
  if tab_ratio ~= 500 then
     table.insert(attributes, {["tabRatio"] = tab_ratio})
  end

  -- Store the firstSheet attribute when it isn't the default.
  if first_sheet > 0 then
     table.insert(attributes, {["firstSheet"] = first_sheet + 1})
  end

  -- Store the activeTab attribute when it isn't the first sheet.
  if active_tab > 0 then
     table.insert(attributes, {["activeTab"] = active_tab})
  end

  self:_xml_empty_tag("workbookView", attributes)
end

----
-- Write <sheets> element.
--
function Workbook:_write_sheets()
  local id_num = 1

  self:_xml_start_tag("sheets")

  for _, worksheet in ipairs(self.worksheets) do
    self:_write_sheet(worksheet.name, id_num, worksheet.hidden)
    id_num = id_num + 1
  end

  self:_xml_end_tag("sheets")
end

----
-- Write <sheet> element.
--
function Workbook:_write_sheet(name, sheet_id, hidden)
  local r_id     = "rId" .. sheet_id

  local attributes = {
    {["name"]    = name},
    {["sheetId"] = sheet_id},
  }

  if hidden then
     table.insert(attributes, {["state"] = "hidden"})
  end

  table.insert(attributes, {["r:id"] = r_id})

  self:_xml_empty_tag("sheet", attributes)
end

----
-- Write <calcPr> element.
--
function Workbook:_write_calc_pr()
  local attributes = {{["calcId"] = 124519}}

  if self.calc_mode == 'manual' then
    table.insert(attributes, {['calcMode'] = self.calc_mode})
    table.insert(attributes, {['calcOnSave'] = "0"})
  elseif self.calc_mode == 'autoNoTable' then
    table.insert(attributes, {['calcMode'] = self.calc_mode})
  end

  if self.calc_on_load then
    table.insert(attributes, {["fullCalcOnLoad"] = "1"})
  end

  self:_xml_empty_tag("calcPr", attributes)
end

----
-- Write the <definedNames> element.
--
function Workbook:_write_defined_names()

  if #self.defined_names == 0 then return end

  self:_xml_start_tag("definedNames")

  for _, defined_name in ipairs(self.defined_names) do
    self:_write_defined_name(aref)
  end

  self:_xml_end_tag("definedNames")
end

----
-- Write the <definedName> element.
--
function Workbook:_write_defined_name(data)
  local name   = data[1]
  local id     = data[2]
  local range  = data[3]
  local hidden = data[4]

  local attributes = {{["name"] = name}}

  if id ~= -1 then
     table.insert(attributes, {["localSheetId"] = id})
  end

  if hidden then
     table.insert(attributes, {["hidden"] = "1"})
  end

  self:_xml_data_element("definedName", range, attributes)
end


return Workbook
