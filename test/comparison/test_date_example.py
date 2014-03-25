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

    Test the date example in the documentation.

    """

    def test_date_example01(self):
        self.run_lua_test('test_date_example01')
