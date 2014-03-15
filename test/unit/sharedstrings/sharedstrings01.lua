----
-- Tests for the xlsxwriter.lua sharedstrings class.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"
require "Test.LongString"

plan(8)

----
-- Tests setup.
--
local expected
local got
local caption
local Sharedstrings = require "xlsxwriter.sharedstrings"
local sharedstrings
local index

-- Remove extra whitespace in the formatted XML strings.
function _clean_xml_string(s)
  return (string.gsub(s, ">%s+<", "><"))
end

----
-- Test the sharedstrings _assemble_xml_file() method.
--
caption = " \tSharedstrings:"
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

index = sharedstrings:_get_string_index("neptune")
is(index, 0, caption .. " _get_string_index()")

index = sharedstrings:_get_string_index("mars")
is(index, 1, caption .. " _get_string_index()")

index = sharedstrings:_get_string_index("venus")
is(index, 2, caption .. " _get_string_index()")

index = sharedstrings:_get_string_index("neptune")
is(index, 0, caption .. " _get_string_index()")

index = sharedstrings:_get_string_index("mars")
is(index, 1, caption .. " _get_string_index()")

index = sharedstrings:_get_string_index("venus")
is(index, 2, caption .. " _get_string_index()")

index = sharedstrings:_get_string_index("venus")
is(index, 2, caption .. " _get_string_index()")


sharedstrings:_assemble_xml_file()

got = _clean_xml_string(sharedstrings:_get_data())

is_string(got, expected, caption .. " _assemble_xml_file()")
