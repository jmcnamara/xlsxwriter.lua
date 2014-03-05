----
-- Tests for the xlsxwriter.lua xml writer class.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"

local Xmlwriter = require "xlsxwriter.xmlwriter"
local writer = Xmlwriter:new()
local expected
local got
local caption

plan(17)

----
-- Test the xml_declaration() method.
--
caption = ' \tTest _xml_declaration()'

writer:_set_filehandle(io.tmpfile())
writer:_xml_declaration()

expected = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_start_tag() with no attributes.
--
caption = ' \tTest _xml_start_tag() with no attributes'

writer:_set_filehandle(io.tmpfile())
writer:_xml_start_tag('foo')

expected = '<foo>'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_start_tag() with attributes.
--
caption = ' \tTest _xml_start_tag() with attributes'

writer:_set_filehandle(io.tmpfile())
writer:_xml_start_tag('foo', {{span = 8}, {baz =7}})

expected = '<foo span="8" baz="7">'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_start_tag() with attributes requiring escaping.
--
caption = ' \tTest _xml_start_tag() with attributes requiring escaping'

writer:_set_filehandle(io.tmpfile())
writer:_xml_start_tag('foo', {{span = '&<>"'}})

expected = '<foo span="&amp;&lt;&gt;&quot;">'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_start_tag_unencoded() with attributes.
--
caption = ' \tTest _xml_start_tag_unencoded() with attributes'

writer:_set_filehandle(io.tmpfile())
writer:_xml_start_tag_unencoded('foo', {{span = '&<>"'}})

expected = '<foo span="&<>"">'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_end_tag().
--
caption = ' \tTest _xml_end_tag()'

writer:_set_filehandle(io.tmpfile())
writer:_xml_end_tag('foo')

expected = '</foo>'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_empty_tag()".
--
caption = ' \tTest _xml_empty_tag()'

writer:_set_filehandle(io.tmpfile())
writer:_xml_empty_tag('foo')

expected = '<foo/>'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_empty_tag() with attributes.
--
caption = ' \tTest _xml_empty_tag() with attributes'

writer:_set_filehandle(io.tmpfile())
writer:_xml_empty_tag('foo', {{span = 8}})

expected = '<foo span="8"/>'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_empty_tag_unencoded() with attributes.
--
caption = ' \tTest _xml_empty_tag_unencoded() with attributes'

writer:_set_filehandle(io.tmpfile())
writer:_xml_empty_tag_unencoded('foo', {{span = '&'}})

expected = '<foo span="&"/>'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_data_element().
--
caption = ' \tTest _xml_data_element()'

writer:_set_filehandle(io.tmpfile())
writer:_xml_data_element('foo', 'bar')

expected = '<foo>bar</foo>'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_data_element() with attributes.
--
caption = ' \tTest _xml_data_element() with attributes'

writer:_set_filehandle(io.tmpfile())
writer:_xml_data_element('foo', 'bar', {{span = '8'}})

expected = '<foo span="8">bar</foo>'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_data_element() with data requiring escaping.
--
caption = ' \tTest _xml_data_element() with data requiring escaping'

writer:_set_filehandle(io.tmpfile())
writer:_xml_data_element('foo', '&<>"', {{span = '8'}})

expected = '<foo span="8">&amp;&lt;&gt;"</foo>'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_string_element().
--
caption = ' \tTest _xml_string_element()'

writer:_set_filehandle(io.tmpfile())
writer:_xml_string_element(99, {{span = 8}})

expected = '<c span="8" t=\"s\"><v>99</v></c>'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_si_element().
--
caption = ' \tTest _xml_si_element()'

writer:_set_filehandle(io.tmpfile())
writer:_xml_si_element('foo', {{span = 8}})

expected = '<si><t span="8">foo</t></si>'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_rich_si_element().
--
caption = ' \tTest _xml_rich_si_element()'

writer:_set_filehandle(io.tmpfile())
writer:_xml_rich_si_element('foo')

expected = '<si>foo</si>'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_number_element().
--
caption = ' \tTest _xml_number_element()'

writer:_set_filehandle(io.tmpfile())
writer:_xml_number_element(99, {{span = 8}})

expected = '<c span="8"><v>99</v></c>'
got = writer:_get_data()

is(got, expected, caption)

----
-- Test _xml_formula_element().
--
caption = ' \tTest _xml_formula_element()'

writer:_set_filehandle(io.tmpfile())
writer:_xml_formula_element('1+2', 3, {{span = 8}})

expected = '<c span="8"><f>1+2</f><v>3</v></c>'
got = writer:_get_data()

is(got, expected, caption)

