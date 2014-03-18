###############################################################################
#
# Base test class for libxlsxwriter functional tests.
#
# Copyright (c), 2014, John McNamara, jmcnamara@cpan.org
#

import unittest
import os
from helper_functions import _compare_xlsx_files


class XLSXBaseTest(unittest.TestCase):
    """
    Test file created with libxlsxwriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        self.no_system_error = 0
        self.got_filename = ''
        self.exp_filename = ''
        self.ignore_files = []
        self.ignore_elements = {}

    def run_lua_test(self, lua_name, exp_filename=None):
        """Run lua test and compare output xlsx file with the Excel file."""

        got = os.system("lua test/comparison/lua/%s.lua" % lua_name)
        self.assertEqual(got, self.no_system_error)

        # Create the path/file names for the xlsx files to compare.
        got_filename = lua_name + '.xlsx'
        if not exp_filename:
            exp_filename = lua_name.replace('test_', '') + '.xlsx'

        self.got_filename = got_filename
        self.exp_filename = 'test/comparison/xlsx_files/' + exp_filename

        # Do the comparison between the files.
        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)
        self.assertEqual(got, exp)

    def tearDown(self):
        # Cleanup.
        if os.path.exists(self.got_filename):
            os.remove(self.got_filename)

        self.ignore_files = []
        self.ignore_elements = {}
