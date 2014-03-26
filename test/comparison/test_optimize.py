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

    Test data writing in optimization, i.e. in-line, mode.

    """

    def test_optimize01(self):
        self.run_lua_test('test_optimize01')

    def test_optimize02(self):
        self.run_lua_test('test_optimize02')
