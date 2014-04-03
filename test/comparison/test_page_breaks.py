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

    Test the worksheet set_x_pagebreaks() methods.

    """

    def test_page_breaks01(self):
        self.ignore_files    = ignore_files
        self.ignore_elements = ignore_elements
        self.run_lua_test('test_page_breaks01')

    def test_page_breaks02(self):
        self.ignore_files    = ignore_files
        self.ignore_elements = ignore_elements
        self.run_lua_test('test_page_breaks02')

    def test_page_breaks03(self):
        self.ignore_files    = ignore_files
        self.ignore_elements = ignore_elements
        self.run_lua_test('test_page_breaks03')

    def test_page_breaks04(self):
        self.ignore_files    = ignore_files
        self.ignore_elements = ignore_elements
        self.run_lua_test('test_page_breaks04')

    def test_page_breaks05(self):
        self.ignore_files    = ignore_files
        self.ignore_elements = ignore_elements
        self.run_lua_test('test_page_breaks05')

    def test_page_breaks06(self):
        self.ignore_files    = ignore_files
        self.ignore_elements = ignore_elements
        self.run_lua_test('test_page_breaks06')
