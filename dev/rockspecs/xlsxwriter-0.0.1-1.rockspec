package = "xlsxwriter"
version = "0.0.1-1"
source = {
  url = "git://github.com/jmcnamara/xlsxwriter.lua.git",
  tag = "v0.0.1"
}
description = {
  summary = "A lua module for creating Excel XLSX files.",
  detailed = [[
      XlsxWriter

      Xlsxwriter is a Lua module that can be used to write text, numbers
      and formulas to multiple worksheets in an Excel 2007+ XLSX file.

      It supports features such as:

        * 100% compatible Excel XLSX files.
        * Full formatting.
        * Memory optimisation mode for writing large files.

      It works with Lua 5.1 and Lua 5.2.
  ]],
  homepage = "http://xlsxwriterlua.readthedocs.org/",
  license = "MIT/X11"
}
dependencies = {
  "lua >= 5.1",
  "zipwriter",
}
build = {
  type = "builtin",
  modules = {
    ["xlsxwriter.app"]           = "xlsxwriter/app.lua",
    ["xlsxwriter.contenttypes"]  = "xlsxwriter/contenttypes.lua",
    ["xlsxwriter.core"]          = "xlsxwriter/core.lua",
    ["xlsxwriter.format"]        = "xlsxwriter/format.lua",
    ["xlsxwriter.packager"]      = "xlsxwriter/packager.lua",
    ["xlsxwriter.relationships"] = "xlsxwriter/relationships.lua",
    ["xlsxwriter.sharedstrings"] = "xlsxwriter/sharedstrings.lua",
    ["xlsxwriter.strict"]        = "xlsxwriter/strict.lua",
    ["xlsxwriter.styles"]        = "xlsxwriter/styles.lua",
    ["xlsxwriter.theme"]         = "xlsxwriter/theme.lua",
    ["xlsxwriter.utility"]       = "xlsxwriter/utility.lua",
    ["xlsxwriter.workbook"]      = "xlsxwriter/workbook.lua",
    ["xlsxwriter.worksheet"]     = "xlsxwriter/worksheet.lua",
    ["xlsxwriter.xmlwriter"]     = "xlsxwriter/xmlwriter.lua",
  }
}
