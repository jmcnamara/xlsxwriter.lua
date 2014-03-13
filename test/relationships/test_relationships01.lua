----
-- Tests for the xlsxwriter.lua relationships class.
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
local Relationships = require "xlsxwriter.relationships"
local relationships

-- Remove extra whitespace in the formatted XML strings.
function _clean_xml_string(s)
  return (string.gsub(s, ">%s+<", "><"))
end

----
-- Test the Relationships  _assemble_xml_file() method.
--
caption = " \tRelationships: Relationships: _assemble_xml_file()"
expected = _clean_xml_string([[
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>]])

relationships = Relationships:new()
relationships:_set_filehandle(io.tmpfile())

relationships:_add_document_relationship('/worksheet', 'worksheets/sheet1.xml')
relationships:_add_document_relationship('/theme', 'theme/theme1.xml')
relationships:_add_document_relationship('/styles', 'styles.xml')
relationships:_add_document_relationship('/sharedStrings', 'sharedStrings.xml')
relationships:_add_document_relationship('/calcChain', 'calcChain.xml')

relationships:_assemble_xml_file()

got = _clean_xml_string(relationships:_get_data())

is_string(got, expected, caption)
