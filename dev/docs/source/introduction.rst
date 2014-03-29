.. _intro:

Introduction
============

**Xlsxwriter** is a Lua module for writing files in the Excel 2007+ XLSX
file format.

It can be used to write text, numbers, and formulas to multiple worksheets and
it supports features such as formatting.

The main advantages of using Xlswriter are:

   * It has a high degree of fidelity with files produced by Excel. In most
     cases the files produced are 100% equivalent to files produced by Excel.

   * It has extensive documentation, example files and tests.

   * It is fast and can be configured to use very little memory even for very
     large output files.

However:

   * It can only create **new files**. It cannot read or modify existing files.

Xlsxwriter is a Lua port of the Perl `Excel::Writer::XLSX <http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX/>`_ and the Python `XlsxWriter <http://xlsxwriter.readthedocs.org>`_ modules and is licensed under an MIT/X11 :ref:`License`.

To try out the module see the next section on :ref:`getting_started`.
