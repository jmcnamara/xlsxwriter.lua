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

    These tests check cell formats.

    """

    def test_format01(self):
        self.run_lua_test('test_format01')

    def test_format02(self):
        self.run_lua_test('test_format02')

    def test_format03(self):
        self.run_lua_test('test_format03')

    # Skip some of the original Perl tests.

    def test_format05(self):
        self.run_lua_test('test_format05')

    def test_format06(self):
        self.run_lua_test('test_format06')

    def test_format07(self):
        self.run_lua_test('test_format07')

    def test_format08(self):
        self.run_lua_test('test_format08')

    def test_format09(self):
        self.run_lua_test('test_format09')

    def test_format10(self):
        self.run_lua_test('test_format10')
