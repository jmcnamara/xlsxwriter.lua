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

    These tests check simple data writing.

    """

    def test_simple01(self):
        self.run_lua_test('test_simple01')

    def test_simple02(self):
        self.run_lua_test('test_simple02')

    def test_simple03(self):
        self.run_lua_test('test_simple03')

    def test_simple04(self):
        self.run_lua_test('test_simple04')

    def test_simple05(self):
        self.run_lua_test('test_simple05')
