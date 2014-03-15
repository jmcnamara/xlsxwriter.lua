----
-- Tests for the xlsxwriter.lua Worksheet class.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"

plan(6)

----
-- Tests setup.
--
local expected
local got
local caption
local Worksheet = require "xlsxwriter.worksheet"
local Format    = require "xlsxwriter.format"
local worksheet
local args = {}


----
-- 1. Test the _write_col_info() method.
--
args["firstcol"]  = 1
args["lastcol"]   = 3
args["width"]     = 5
args["format"]    = false
args["hidden"]    = false
args["level"]     = 0
args["collapsed"] = false

caption  = " \tWorksheet: _write_col_info(args)"
expected = '<col min="2" max="4" width="5.7109375" customWidth="1"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:_write_col_info(args)

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 2. Test the _write_col_info() method.
--
args["firstcol"]  = 5
args["lastcol"]   = 5
args["width"]     = 8
args["format"]    = false
args["hidden"]    = true
args["level"]     = 0
args["collapsed"] = false

caption  = " \tWorksheet: _write_col_info(args)"
expected = '<col min="6" max="6" width="8.7109375" hidden="1" customWidth="1"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:_write_col_info(args)

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 3. Test the _write_col_info() method.
--
args["firstcol"]  = 7
args["lastcol"]   = 7
args["width"]     = nil
args["format"]    = Format:new{xf_index = 1}
args["hidden"]    = false
args["level"]     = 0
args["collapsed"] = false

caption  = " \tWorksheet: _write_col_info(args)"
expected = '<col min="8" max="8" width="9.140625" style="1"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:_write_col_info(args)

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 4. Test the _write_col_info() method.
--
args["firstcol"]  = 8
args["lastcol"]   = 8
args["width"]     = 8.43
args["format"]    = Format:new{xf_index = 1}
args["hidden"]    = false
args["level"]     = 0
args["collapsed"] = false

caption  = " \tWorksheet: _write_col_info(args)"
expected = '<col min="9" max="9" width="9.140625" style="1"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:_write_col_info(args)

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 5. Test the _write_col_info() method.
--
args["firstcol"]  = 9
args["lastcol"]   = 9
args["width"]     = 2
args["format"]    = false
args["hidden"]    = false
args["level"]     = 0
args["collapsed"] = false

caption  = " \tWorksheet: _write_col_info(args)"
expected = '<col min="10" max="10" width="2.7109375" customWidth="1"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())
worksheet:_write_col_info(args)

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 6. Test the _write_col_info() method.
--
args["firstcol"]  = 11
args["lastcol"]   = 11
args["width"]     = nil
args["format"]    = false
args["hidden"]    = true
args["level"]     = 0
args["collapsed"] = false

caption  = " \tWorksheet: _write_col_info(args)"
expected = '<col min="12" max="12" width="0" hidden="1" customWidth="1"/>'

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:_write_col_info(args)

got = worksheet:_get_data()

is(got, expected, caption)
