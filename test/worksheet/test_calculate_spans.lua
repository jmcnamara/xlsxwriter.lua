----
-- Tests for the xlsxwriter.lua Worksheet class.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"

plan(18)

----
-- Tests setup.
--
local got
local expected
local caption
local row
local col
local Worksheet = require "xlsxwriter.worksheet"
local worksheet

----
-- 1. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 0
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:16', [1] = '17:17'}

is_deeply(got, expected, caption)

----
-- 2. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 1
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:15', [1] = '16:17'}

is_deeply(got, expected, caption)

----
-- 3. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 2
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:14', [1] = '15:17'}

is_deeply(got, expected, caption)

----
-- 4. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 3
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:13', [1] = '14:17'}

is_deeply(got, expected, caption)

----
-- 5. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 4
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:12', [1] = '13:17'}

is_deeply(got, expected, caption)

----
-- 6. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 5
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:11', [1] = '12:17'}

is_deeply(got, expected, caption)

----
-- 7. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 6
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:10', [1] = '11:17'}

is_deeply(got, expected, caption)

----
-- 8. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 7
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:9', [1] = '10:17'}

is_deeply(got, expected, caption)

----
-- 9. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 8
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:8', [1] = '9:17'}

is_deeply(got, expected, caption)

----
-- 10. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 9
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:7', [1] = '8:17'}

is_deeply(got, expected, caption)

----
-- 11. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 10
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:6', [1] = '7:17'}

is_deeply(got, expected, caption)

----
-- 12. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 11
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:5', [1] = '6:17'}

is_deeply(got, expected, caption)

----
-- 13. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 12
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:4', [1] = '5:17'}

is_deeply(got, expected, caption)

----
-- 14. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 13
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:3', [1] = '4:17'}

is_deeply(got, expected, caption)

----
-- 15. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 14
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:2', [1] = '3:17'}

is_deeply(got, expected, caption)

----
-- 16. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 15
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[0] = '1:1', [1] = '2:17'}

is_deeply(got, expected, caption)

----
-- 17. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 16
col       = 0
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[1] = '1:16', [2] = '17:17'}

is_deeply(got, expected, caption)

----
-- 18. Test _calculate_spans() method for range (row, col), (row+16, col+16).
--
row       = 16
col       = 1
caption   = " \tWorksheet: _calculate_spans()"
worksheet = Worksheet:new()

for i = row, row + 16 do
  worksheet:write(i, col, 1)
  col = col + 1
end

worksheet:_calculate_spans()

got = worksheet.row_spans
expected = {[1] = '2:17', [2] = '18:18'}

is_deeply(got, expected, caption)

