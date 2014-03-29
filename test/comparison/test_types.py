###############################################################################
#
# Test cases for xlsxwriter.lua.
#
# Copyright (c), 2014, John McNamara, jmcnamara@cpan.org
#

import base_test_class

class TestCompareXLSXFiles(base_test_class.XLSXBaseTest):
    """
    Test a file created with xlsxwriter.lua against a file created by Excel.

    Test the conversion of Lua types to Excel types.

    """

    def test_types02(self):
        self.run_lua_test('test_types02')

    def test_types12(self):
        self.run_lua_test('test_types12', 'types02.xlsx')
