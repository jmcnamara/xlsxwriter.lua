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

    Tests for XML and URL escaping.

    """

    def test_escapes01(self):
        self.ignore_files = ['xl/calcChain.xml',
                             '[Content_Types].xml',
                             'xl/_rels/workbook.xml.rels']
        self.run_lua_test('test_escapes01')

    def test_escapes04(self):
        self.run_lua_test('test_escapes04')

    def test_escapes05(self):
        self.run_lua_test('test_escapes05')

    def test_escapes06(self):
        self.run_lua_test('test_escapes06')

    def test_escapes07(self):
        self.run_lua_test('test_escapes07')

    def test_escapes08(self):
        self.run_lua_test('test_escapes08')
