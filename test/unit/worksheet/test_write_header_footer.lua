----
-- Tests for the xlsxwriter.lua.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"

plan(5)

----
-- Tests setup.
--
local expected
local got
local caption
local Worksheet = require 'xlsxwriter.worksheet'
local worksheet

----
-- Test the _write_odd_header() method.
--
caption  = " \tWorksheet: _write_odd_header()"
expected = '<oddHeader>Page &amp;P of &amp;N</oddHeader>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:set_header('Page &P of &N')

worksheet:_write_odd_header()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_odd_footer() method.
--
caption  = " \tWorksheet: _write_odd_footer()"
expected = '<oddFooter>&amp;F</oddFooter>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:set_footer('&F')

worksheet:_write_odd_footer()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_header_footer() method. Header only.
--
caption  = " \tWorksheet: _write_header_footer()"
expected = '<headerFooter><oddHeader>Page &amp;P of &amp;N</oddHeader></headerFooter>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:set_header('Page &P of &N')

worksheet:_write_header_footer()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_header_footer() method. Footer only.
--
caption  = " \tWorksheet: _write_header_footer()"
expected = '<headerFooter><oddFooter>&amp;F</oddFooter></headerFooter>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:set_footer('&F')

worksheet:_write_header_footer()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- Test the _write_header_footer() method. Header and footer.
--
caption  = " \tWorksheet: _write_header_footer()"
expected = '<headerFooter><oddHeader>Page &amp;P of &amp;N</oddHeader><oddFooter>&amp;F</oddFooter></headerFooter>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:set_header('Page &P of &N')
worksheet:set_footer('&F')

worksheet:_write_header_footer()

got = worksheet:_get_data()

is(got, expected, caption)

