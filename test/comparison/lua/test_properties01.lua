----
-- Test cases for xlsxwriter.lua.
--
-- Test workbook properties.
--
-- Copyright 2014, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local workbook  = Workbook:new("test_properties01.xlsx")
local worksheet = workbook:add_worksheet()

workbook:set_properties({
    title    = "This is an example spreadsheet",
    subject  = "With document properties",
    author   = "Someone",
    manager  = "Dr. Heinz Doofenshmirtz",
    company  = "of Wolves",
    category = "Example spreadsheets",
    keywords = "Sample, Example, Properties",
    comments = "Created with Perl and Excel::Writer::XLSX",
    status   = "Quo",
})

worksheet:set_column("A:A", 70)
worksheet:write("A1", "Select 'Office Button -> Prepare -> Properties' to see the file properties.")

workbook:close()
