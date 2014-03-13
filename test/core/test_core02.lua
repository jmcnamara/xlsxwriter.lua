----
-- Tests for the xlsxwriter.lua core class.
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
local Core = require "xlsxwriter.core"
local core

-- Remove extra whitespace in the formatted XML strings.
function _clean_xml_string(s)
  return (string.gsub(s, ">%s+<", "><"))
end

----
-- Test the Core  _assemble_xml_file() method.
--
caption = " \tCore: Core: _assemble_xml_file()"
expected = _clean_xml_string([[
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>This is an example spreadsheet</dc:title>
  <dc:subject>With document properties</dc:subject>
  <dc:creator>John McNamara</dc:creator>
  <cp:keywords>Sample, Example, Properties</cp:keywords>
  <dc:description>Created with Perl and Excel::Writer::XLSX</dc:description>
  <cp:lastModifiedBy>John McNamara</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2011-04-06T19:45:15Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2011-04-06T19:45:15Z</dcterms:modified>
  <cp:category>Example spreadsheets</cp:category>
  <cp:contentStatus>Quo</cp:contentStatus>
</cp:coreProperties>]])

core = Core:new()
core:_set_filehandle(io.tmpfile())

core:_set_properties{
  ["title"]    = 'This is an example spreadsheet',
  ["subject"]  = 'With document properties',
  ["author"]   = 'John McNamara',
  ["manager"]  = 'Dr. Heinz Doofenshmirtz',
  ["company"]  = 'of Wolves',
  ["category"] = 'Example spreadsheets',
  ["keywords"] = 'Sample, Example, Properties',
  ["comments"] = 'Created with Perl and Excel::Writer::XLSX',
  ["status"]   = 'Quo',
  ["created"]  = os.date("%Y-%m-%dT%H:%M:%SZ",
                         os.time{year=2011, month=4, day=6,
                                 hour=19, min=45, sec=15})
}

core:_assemble_xml_file()

got = _clean_xml_string(core:_get_data())

is_string(got, expected, caption)
