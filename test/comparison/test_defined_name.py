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

    Test defined names in the workbook.

    """

    def test_defined_name02(self):
        self.run_lua_test('test_defined_name02')

    def test_defined_name03(self):
        self.run_lua_test('test_defined_name03')

    def test_defined_name04(self):
        self.run_lua_test('test_defined_name04')
