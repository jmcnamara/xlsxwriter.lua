###############################################################################
#
# Test cases for xlsxwriter.lua.
#
# Copyright (c), 2014, John McNamara, jmcnamara@cpan.org
#

import base_test_class

ignore_files    = ['xl/printerSettings/printerSettings1.bin',
                   'xl/worksheets/_rels/sheet1.xml.rels']
ignore_elements = {'[Content_Types].xml': ['<Default Extension="bin"'],
                   'xl/worksheets/sheet1.xml': ['<pageMargins', '<pageSetup']}

class TestCompareXLSXFiles(base_test_class.XLSXBaseTest):
    """
    Test file created with xlsxwriter.lua against a file created by Excel.

    Test setting the first  worksheet page to be printed.

    """

    def test_set_start_page01(self):
        self.ignore_files    = ignore_files
        self.ignore_elements = ignore_elements
        self.run_lua_test('test_set_start_page01')
