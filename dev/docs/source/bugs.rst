.. highlight:: lua

.. _bugs:

Known Issues and Bugs
=====================

This section lists known issues and bugs and gives some information on how to
submit bug reports.

Content is Unreadable. Open and Repair
--------------------------------------

Very, very occasionally you may see an Excel warning when opening an ``xlsxwriter`` file like:

   Excel could not open file.xlsx because some content is unreadable. Do you
   want to open and repair this workbook.

This ominous sounding message is Excel's default warning for any validation
error in the XML used for the components of the XLSX file.

If you encounter an issue like this you should open an issue on GitHub with a
program to replicate the issue (see below) or send one of the failing output
files to the :ref:`author`.


Formulas displayed as ``#NAME?`` until edited
---------------------------------------------

Excel 2010 and 2013 added functions which weren't defined in the original file
specification. These functions are referred to as *future* functions. Examples
of these functions are ``ACOT``, ``CHISQ.DIST.RT`` , ``CONFIDENCE.NORM``,
``STDEV.P``, ``STDEV.S`` and ``WORKDAY.INTL``. The full list is given in the
`MS XLSX extensions documentation on future functions <http://msdn.microsoft.com/en-us/library/dd907480%28v=office.12%29.aspx>`_.

When written using ``write_formula()`` these functions need to be fully
qualified with the ``_xlfn.`` prefix as they are shown in the MS XLSX
documentation link above. For example::

    worksheet:write_formula('A1', '=_xlfn.STDEV.S(B1:B10)')


Formula results displaying as zero in non-Excel applications
------------------------------------------------------------

Due to wide range of possible formulas and interdependencies between them,
``xlsxwriter`` doesn't, and realistically cannot, calculate the result of a
formula when it is written to an XLSX file. Instead, it stores the value 0 as
the formula result. It then sets a global flag in the XLSX file to say that
all formulas and functions should be recalculated when the file is opened.

This is the method recommended in the Excel documentation and in general it
works fine with spreadsheet applications. However, applications that don't
have a facility to calculate formulas, such as Excel Viewer, or several mobile
applications, will only display the 0 results.

If required, it is also possible to specify the calculated result of the
formula using the optional ``value`` parameter in :func:`write_formula()`::

    worksheet:write_formula('A1', '=2+2', num_format, 4)


Strings aren't displayed in Apple Numbers in 'constant_memory' mode
-------------------------------------------------------------------

In :func:`Workbook` ``'constant_memory'`` mode ``xlsxwriter`` uses an optimisation where cell strings aren't stored in an Excel structure call "shared strings"
and instead are written "in-line".

This is a documented Excel feature that is supported by most spreadsheet
applications. One known exception is Apple Numbers for Mac where the string
data isn't displayed.


Images not displayed correctly in Excel 2001 for Mac and non-Excel applications
-------------------------------------------------------------------------------

Images inserted into worksheets via :func:`insert_image` may not display
correctly in Excel 2011 for Mac and non-Excel applications such as OpenOffice
and LibreOffice. Specifically the images may looked stretched or squashed.

This is not specifically an ``xlsxwriter`` issue. It also occurs with files created in Excel 2007 and Excel 2010.



Reporting Bugs
==============

Here are some tips on reporting bugs in ``xlsxwriter``.


Upgrade to the latest version of the module
-------------------------------------------

The bug you are reporting may already be fixed in the latest version of the
module. You can check which version of ``xlsxwriter`` that you are using as
follows::

    lua -e 'W = require "xlsxwriter.workbook"; print(W.version)'

Check the :ref:`changes` section to see what has changed in the latest versions.


Read the documentation
----------------------

Read or search the ``xlsxwriter`` documentation to see if the issue you are
encountering is already explained.

Look at the example programs
----------------------------

There are many :ref:`example_programs` in the distribution. Try to identify an example
program that corresponds to your query and adapt it to use as a bug report.

Use the xlsxwriter Issue tracker on GitHub
------------------------------------------

The `xlsxwriter issue tracker <https://github.com/jmcnamara/xlsxwriter.lua/issues>`_ is on GitHub.


Pointers for submitting a bug report
------------------------------------

#. Describe the problem as clearly and as concisely as possible.

#. Include a sample program. This is probably the most important step. It is
   generally easier to describe a problem in code than in written prose.

#. The sample program should be as small as possible to demonstrate the
   problem. Don't copy and paste large non-relevant sections of your program.

A sample bug report is shown below. This format helps analyse and respond to
the bug report more quickly.

   **Issue with SOMETHING**

   I am using xlsxwriter to do SOMETHING but it appears to do SOMETHING ELSE.

   I am using Lua version X.Y and xlsxwriter x.y.z.

   Here is some code that demonstrates the problem::

     local Workbook = require "xlsxwriter.workbook"

     local workbook  = Workbook:new("hello_world.xlsx")
     local worksheet = workbook:add_worksheet()

     worksheet:write("A1", "Hello world")

     workbook:close()
