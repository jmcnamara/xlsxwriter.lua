----
-- Tests for the xlsxwriter.lua.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"

plan(18)

----
-- Tests setup.
--
local expected
local got
local caption
local Worksheet = require 'xlsxwriter.worksheet'
local worksheet
local password
local options

----
-- 1. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" objects="1" scenarios="1"/>'

password = ''
options  = {}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 2. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection password="83AF" sheet="1" objects="1" scenarios="1"/>'

password = 'password'
options  = {}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 3. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" objects="1" scenarios="1" selectLockedCells="1"/>'

password = ''
options  = {["select_locked_cells"] = false}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 4. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0"/>'

password = ''
options  = {["format_cells"] = true}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 5. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" objects="1" scenarios="1" formatColumns="0"/>'

password = ''
options  = {["format_columns"] = true}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 6. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" objects="1" scenarios="1" formatRows="0"/>'

password = ''
options  = {["format_rows"] = true}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 7. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" objects="1" scenarios="1" insertColumns="0"/>'

password = ''
options  = {["insert_columns"] = true}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 8. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" objects="1" scenarios="1" insertRows="0"/>'

password = ''
options  = {["insert_rows"] = true}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 9. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" objects="1" scenarios="1" insertHyperlinks="0"/>'

password = ''
options  = {["insert_hyperlinks"] = true}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 10. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" objects="1" scenarios="1" deleteColumns="0"/>'

password = ''
options  = {["delete_columns"] = true}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 11. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" objects="1" scenarios="1" deleteRows="0"/>'

password = ''
options  = {["delete_rows"] = true}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 12. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" objects="1" scenarios="1" sort="0"/>'

password = ''
options  = {["sort"] = true}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 13. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" objects="1" scenarios="1" autoFilter="0"/>'

password = ''
options  = {["autofilter"] = true}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 14. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" objects="1" scenarios="1" pivotTables="0"/>'

password = ''
options  = {["pivot_tables"] = true}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 15. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" scenarios="1"/>'

password = ''
options  = {["objects"] = true}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 16. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" objects="1"/>'

password = ''
options  = {["scenarios"] = true}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 17. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0" selectLockedCells="1" selectUnlockedCells="1"/>'

password = ''
options  = {
  ["format_cells"]          = true,
  ["select_locked_cells"]   = false,
  ["select_unlocked_cells"] = false
}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)

----
-- 18. Test the _write_sheet_protection() method.
--
caption  = " \tWorksheet: _write_sheet_protection()"
expected = '<sheetProtection password="996B" sheet="1" formatCells="0" formatColumns="0" formatRows="0" insertColumns="0" insertRows="0" insertHyperlinks="0" deleteColumns="0" deleteRows="0" selectLockedCells="1" sort="0" autoFilter="0" pivotTables="0" selectUnlockedCells="1"/>'

password = 'drowssap'
options  = {
  ["objects"]               = true,
  ["scenarios"]             = true,
  ["format_cells"]          = true,
  ["format_columns"]        = true,
  ["format_rows"]           = true,
  ["insert_columns"]        = true,
  ["insert_rows"]           = true,
  ["insert_hyperlinks"]     = true,
  ["delete_columns"]        = true,
  ["delete_rows"]           = true,
  ["select_locked_cells"]   = false,
  ["sort"]                  = true,
  ["autofilter"]            = true,
  ["pivot_tables"]          = true,
  ["select_unlocked_cells"] = false,
}

worksheet = Worksheet:new()
worksheet:_set_filehandle(io.tmpfile())

worksheet:protect(password, options)
worksheet:_write_sheet_protection()

got = worksheet:_get_data()

is(got, expected, caption)
