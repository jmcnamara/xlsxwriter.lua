.. _ex_utf8:

Example: Write UTF-8 Strings
============================

An example of writing simple UTF-8 strings to a worksheet.

Unicode strings in Excel must be UTF-8 encoded. With ``xlsxwriter`` all that
is required is that the source file is UTF-8 encoded and Lua will handle the
UTF-8 strings like any other strings:

.. image:: _images/utf8.png


.. only:: html

   .. literalinclude:: ../../../examples/utf8.lua
      :language: lua
