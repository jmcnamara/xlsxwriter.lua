###############################################################################
#
# Tests for libxlsxwriter.
#
# Copyright (c), 2014, John McNamara, jmcnamara@cpan.org
#

import base_test_class

class TestCompareXLSXFiles(base_test_class.XLSXBaseTest):
    """
    Test file created with libxlsxwriter against a file created by Excel.

    """

    def test_data01(self):
        self.run_lua_test('test_data01')

    def test_data01(self):
        self.run_lua_test('test_data02')
