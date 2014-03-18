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

    These tests check date writing functions.

    """

    def test_date_1904_01(self):
        self.run_lua_test('test_date_1904_01')

    def test_date_1904_02(self):
        self.run_lua_test('test_date_1904_02')

    def test_date_1904_03(self):
        self.run_lua_test('test_date_1904_03', 'date_1904_01.xlsx')

    def test_date_1904_04(self):
        self.run_lua_test('test_date_1904_04', 'date_1904_02.xlsx')
