----
-- Tests for the xlsxwriter.lua contenttypes class.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"
require "Test.LongString"

plan(1)

----
-- Tests setup.
--
local expected
local got
local caption
local Contenttypes = require "xlsxwriter.contenttypes"
local contenttypes

-- Remove extra whitespace in the formatted XML strings.
function _clean_xml_string(s)
  return (string.gsub(s, ">%s+<", "><"))
end

----
-- Test the Contenttypes  _assemble_xml_file() method.
--
caption = " \tContenttypes: Contenttypes: _assemble_xml_file()"
expected = _clean_xml_string([[
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">

  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>

  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>]])

contenttypes = Contenttypes:new()
contenttypes:_set_filehandle(io.tmpfile())

contenttypes:_add_worksheet_name('sheet1')
contenttypes:_add_default('jpeg', 'image/jpeg')
contenttypes:_add_shared_strings()
contenttypes:_add_calc_chain()

contenttypes:_assemble_xml_file()

got = _clean_xml_string(contenttypes:_get_data())

is_string(got, expected, caption)
