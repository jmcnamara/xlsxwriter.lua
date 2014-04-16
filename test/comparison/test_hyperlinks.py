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
