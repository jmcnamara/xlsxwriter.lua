###############################################################################
#
# Test cases for xlsxwriter.lua.
#
# Copyright (c), 2014, John McNamara, jmcnamara@cpan.org
#

import base_test_class

class TestCompareXLSXFiles(base_test_class.XLSXBaseTest):
    """
    Test file created with xlsxwriter.lua against a file created by Excel.

    Tests for hyperlinks in files.

    """

    def test_hyperlink01(self):
        self.run_lua_test('test_hyperlink01')

    def test_hyperlink02(self):
        self.run_lua_test('test_hyperlink02')

    def test_hyperlink03(self):
        self.run_lua_test('test_hyperlink03')

    def test_hyperlink04(self):
        self.run_lua_test('test_hyperlink04')

    def test_hyperlink05(self):
        self.run_lua_test('test_hyperlink05')

    def test_hyperlink06(self):
        self.run_lua_test('test_hyperlink06')

    def test_hyperlink07(self):
        self.run_lua_test('test_hyperlink07')

    def test_hyperlink08(self):
        self.run_lua_test('test_hyperlink08')

    def test_hyperlink09(self):
        self.run_lua_test('test_hyperlink09')

    def test_hyperlink10(self):
        self.run_lua_test('test_hyperlink10')

    def test_hyperlink11(self):
        self.run_lua_test('test_hyperlink11')

    def test_hyperlink12(self):
        self.run_lua_test('test_hyperlink12')

    def test_hyperlink13(self):
        self.run_lua_test('test_hyperlink13')

    def test_hyperlink14(self):
        self.run_lua_test('test_hyperlink14')

    def test_hyperlink15(self):
        self.run_lua_test('test_hyperlink15')

    def test_hyperlink16(self):
        self.run_lua_test('test_hyperlink16')

    def test_hyperlink17(self):
        self.run_lua_test('test_hyperlink17')

    def test_hyperlink18(self):
        self.run_lua_test('test_hyperlink18')

    def test_hyperlink19(self):
        self.ignore_files = ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels']
        self.run_lua_test('test_hyperlink19')

    def test_hyperlink20(self):
        self.run_lua_test('test_hyperlink20')
