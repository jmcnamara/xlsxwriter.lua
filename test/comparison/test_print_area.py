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

    Test the worksheet print area.

    """

    def test_print_area01(self):
        self.ignore_files    = ignore_files
        self.ignore_elements = ignore_elements
        self.run_lua_test('test_print_area01')

    def test_print_area02(self):
        self.ignore_files    = ignore_files
        self.ignore_elements = ignore_elements
        self.run_lua_test('test_print_area02')

    def test_print_area03(self):
        self.ignore_files    = ignore_files
        self.ignore_elements = ignore_elements
        self.run_lua_test('test_print_area03')

    def test_print_area04(self):
        self.ignore_files    = ignore_files
        self.ignore_elements = ignore_elements
        self.run_lua_test('test_print_area04')

    def test_print_area05(self):
        self.ignore_files    = ignore_files
        self.ignore_elements = ignore_elements
        self.run_lua_test('test_print_area05')

    def test_print_area06(self):
        self.ignore_files    = ignore_files
        self.ignore_elements = ignore_elements
        self.run_lua_test('test_print_area06')

    def test_print_area07(self):
        self.ignore_files    = ignore_files
        self.ignore_elements = ignore_elements
        self.run_lua_test('test_print_area07')
