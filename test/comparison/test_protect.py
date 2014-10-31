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

    These tests check cell protection.

    """

    def test_protect01(self):
        self.run_lua_test('test_protect01')

    def test_protect02(self):
        self.run_lua_test('test_protect02')

    def test_protect03(self):
        self.run_lua_test('test_protect03')
