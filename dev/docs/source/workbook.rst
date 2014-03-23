.. highlight:: lua

.. _workbook:

The Workbook Class
==================

The Workbook class is the main class exposed by the xlsxwriter.lua module and it is
the only class that you will need to instantiate directly.

The Workbook class represents the entire spreadsheet as you see it in Excel and
internally it represents the Excel file as it is written on disk.


.. _constructor:

Constructor
-----------

.. function:: Workbook:new(filename[,options])

   Create a new xlsxwriter.lua Workbook object.

   :param filename: The name of the new Excel file to create.
   :param options:  Optional workbook parameters. See below.
   :rtype:          A Workbook object.


The ``Workbook:new()`` constructor is used to create a new Excel workbook with a
given filename::

    local Workbook = require "xlsxwriter.workbook"

    workbook  = Workbook:new("filename.xlsx")
    worksheet = workbook:add_worksheet()

    worksheet:write(0, 0, "Hello Excel")

    workbook:close()

.. image:: _images/workbook01.png

The constructor options are:

* **constant_memory**: Reduces the amount of data stored in memory so that
  large files can be written efficiently::

       workbook = Workbook:new(filename, {constant_memory = true})

  Note, in this mode a row of data is written and then discarded when a cell
  in a new row is added via one of the worksheet ``write_()`` methods.
  Therefore, once this mode is active, data should be written in sequential
  row order.

  See :ref:`memory_perf` for more details.

* **strings_to_numbers**: Enable the
  :ref:`worksheet: <Worksheet>`:func:`write()` method to convert strings to
  numbers, where possible, using ``tonumber()`` in order to avoid an Excel
  warning about "Numbers Stored as Text". The default is ``false``::

      workbook = Workbook:new(filename, {strings_to_numbers = true})

* **strings_to_formulas**: Enable the
  :ref:`worksheet: <Worksheet>`:func:`write()` method to convert strings to
  formulas. The default is ``true``::

      workbook = Workbook:new(filename, {strings_to_formulas = false})

* **default_date_format**: This option is used to specify a default date
  format string for use with the
  :ref:`worksheet: <Worksheet>`:func:`write_datetime()` method when an
  explicit format isn't given. See :ref:`working_with_dates_and_time` for more
  details::

      Workbook:new(filename, {default_date_format = "dd/mm/yy"})

* **date_1904**: Excel for Windows uses a default epoch of 1900 and Excel for
  Mac uses an epoch of 1904. However, Excel on either platform will convert
  automatically between one system and the other. xlsxwriter.lua stores dates in
  the 1900 format by default. If you wish to change this you can use the
  ``date_1904`` workbook option. This option is mainly for enhanced
  compatibility with Excel and in general isn't required very often::

      workbook = Workbook:new(filename, {date_1904 = true})

When specifying a filename it is recommended that you use an ``.xlsx``
extension or Excel will generate a warning when opening the file.


workbook:add_worksheet()
------------------------

.. function:: add_worksheet([sheetname])

   Add a new worksheet to a workbook:

   :param sheetname: Optional worksheet name, defaults to Sheet1, etc.
   :rtype: A :ref:`worksheet <Worksheet>` object.

The ``add_worksheet()`` method adds a new worksheet to a workbook.

At least one worksheet should be added to a new workbook. The
:ref:`Worksheet <worksheet>` object is used to write data and configure a
worksheet in the workbook.

The ``sheetname`` parameter is optional. If it is not specified the default
Excel convention will be followed, i.e. Sheet1, Sheet2, etc.::

    worksheet1 = workbook:add_worksheet()          -- Sheet1
    worksheet2 = workbook:add_worksheet("Foglio2") -- Foglio2
    worksheet3 = workbook:add_worksheet("Data")    -- Data
    worksheet4 = workbook:add_worksheet()          -- Sheet4

.. image:: _images/workbook02.png

The worksheet name must be a valid Excel worksheet name, i.e. it cannot contain
any of the characters ``[ ] : * ? / \`` and it must be less than 32
characters.

In addition, you cannot use the same, case insensitive, ``sheetname`` for more
than one worksheet.

workbook:add_format()
---------------------

.. function:: add_format([properties])

   Create a new Format object to formats cells in worksheets.

   :paramionary properties: An optional table of format properties.
   :rtype: A :ref:`Format <Format>` object.

The ``add_format()`` method can be used to create new :ref:`Format <Format>`
objects which are used to apply formatting to a cell. You can either define
the properties at creation time via a table of property values or later
via method calls::

    format1 = workbook:add_format(props) -- Set properties at creation.
    format2 = workbook:add_format()      -- Set properties later.

See the :ref:`format` and :ref:`working_with_formats` sections for more details
about Format properties and how to set them.


workbook:close()
----------------

.. function:: close()

   Close the Workbook object and write the XLSX file.

This should be done for every file.

    workbook:close()

Currently, there is no implicit close().
