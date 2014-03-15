----
-- Tests for the xlsxwriter.lua xml writer class.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

require "Test.More"

local utility = require "xlsxwriter.utility"
local expected
local got
local caption
local color

plan(4)

----
-- Test color conversions.
--
color    = 'blue'
caption  = string.format(" \tcolor('%s')", color)
expected = 'FF0000FF'
got      = utility.excel_color(color)

is(got, expected, caption)

----
-- Test color conversions.
--
color    = 'yellow'
caption  = string.format(" \tcolor('%s')", color)
expected = 'FFFFFF00'
got      = utility.excel_color(color)

is(got, expected, caption)

----
-- Test color conversions.
--
color    = '#0000FF'
caption  = string.format(" \tcolor('%s')", color)
expected = 'FF0000FF'
got      = utility.excel_color(color)

is(got, expected, caption)

----
-- Test color conversions.
--
color    = '#0000Ff'
caption  = string.format(" \tcolor('%s')", color)
expected = 'FF0000FF'
got      = utility.excel_color(color)

is(got, expected, caption)
