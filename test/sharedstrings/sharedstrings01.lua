----
-- Tests for the xlsxwriter.lua sharedstrings class.
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
local Sharedstrings = require "xlsxwriter.sharedstrings"
local sharedstrings

-- Remove extra whitespace in the formatted XML strings.
function _clean_xml_string(s)
  return (string.gsub(s, ">%s+<", "><"))
end

----
-- Test the sharedstrings _assemble_xml_file() method.
--
caption = " \tSharedstrings: Sharedstrings: _assemble_xml_file()"
expected = _clean_xml_string([[
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="7" uniqueCount="3">
  <si>
    <t>neptune</t>
  </si>
  <si>
    <t>mars</t>
  </si>
  <si>
    <t>venus</t>
  </si>
</sst>]])

sharedstrings = Sharedstrings:new()
sharedstrings:_set_filehandle(io.tmpfile())

sharedstrings.string_count = 7
sharedstrings.unique_count = 3
sharedstrings.string_array = {'neptune', 'mars', 'venus'}

sharedstrings:_assemble_xml_file()

got = _clean_xml_string(sharedstrings:_get_data())

is_string(got, expected, caption)
