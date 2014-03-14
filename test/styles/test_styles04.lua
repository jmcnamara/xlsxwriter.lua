----
-- Tests for the xlsxwriter.lua styles class.
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
local Styles   = require "xlsxwriter.styles"
local Workbook = require "xlsxwriter.workbook"
local styles
local workbook

-- Remove extra whitespace in the formatted XML strings.
function _clean_xml_string(s)
  return (string.gsub(s, ">%s+<", "><"))
end

----
-- Test the _write_styles() method.
--
caption = " \tStyles: Styles: _assemble_xml_file()"
expected = _clean_xml_string([[
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="2">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="gray125"/>
    </fill>
  </fills>
  <borders count="14">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="hair">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="dotted">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="dashDotDot">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="dashDot">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="dashed">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="thin">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="mediumDashDotDot">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="slantDashDot">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="mediumDashDot">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="mediumDashed">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="medium">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="thick">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="double">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="14">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="5" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="6" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="7" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="8" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="9" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="10" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="11" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="12" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="13" xfId="0" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>]])

styles   = Styles:new()
workbook = Workbook:new("test.xlsx")

workbook:add_format{top = 7 }
workbook:add_format{top = 4 }
workbook:add_format{top = 11}
workbook:add_format{top = 9 }
workbook:add_format{top = 3 }
workbook:add_format{top = 1 }
workbook:add_format{top = 12}
workbook:add_format{top = 13}
workbook:add_format{top = 10}
workbook:add_format{top = 8 }
workbook:add_format{top = 2 }
workbook:add_format{top = 5 }
workbook:add_format{top = 6 }

workbook:_set_default_xf_indices()
workbook:_prepare_format_properties()

local properties = {}
properties["xf_formats"]       = workbook.xf_formats
properties["palette"]          = workbook.palette
properties["font_count"]       = workbook.font_count
properties["num_format_count"] = workbook.num_format_count
properties["border_count"]     = workbook.border_count
properties["fill_count"]       = workbook.fill_count
properties["custom_colors"]    = workbook.custom_colors
properties["dxf_formats"]      = workbook.dxf_formats

styles:_set_style_properties(properties)

styles:_set_filehandle(io.tmpfile())

styles:_assemble_xml_file()

got = _clean_xml_string(styles:_get_data())

is_string(got, expected, caption)
