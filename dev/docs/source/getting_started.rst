.. _getting_started:

Getting Started with xlsxwriter
===============================

Here are some easy instructions to get you up and running with the xlsxwriter
module.


Installing xlsxwriter
---------------------

Xlsxwriter is a pure Lua module and doesn't need a native compiler to install. However, it has a dependency
on the `ZipWriter <https://github.com/moteus/ZipWriter>`_ module which does have binary dependencies.

These dependencies are handled automatically if you use the ``luarocks`` or ``luadist`` methods shown below.


Using luarocks
**************

The easiest way to install ``xlsxwriter`` is with the `luarocks <http://luarocks.org>`_ utility::

    $ sudo luarocks install xlsxwriter


Using luadist
*************

Another easy "packaged" way of installing ``xlsxwriter`` is with the `luadist <http://luadist.org>`_ distribution::

    $ sudo luadist install xlsxwriter


Cloning from GitHub
*******************

The ``xlsxwriter`` source code and bug tracker is in the
`xlsxwriter.lua repository <http://github.com/jmcnamara/xlsxwriter.lua>`_ on GitHub.
You can clone the repository and install from it as follows::

    $ git clone https://github.com/jmcnamara/xlsxwriter.lua.git

    $ cd xlsxwriter.lua
    $ sudo luarocks make
    # or
    $ sudo luadist make

Running a sample program
------------------------

If the installation went correctly you can create a small sample program like
the following to verify that the module works correctly:

.. code-block:: lua

    local Workbook = require "xlsxwriter.workbook"

    local workbook  = Workbook:new("hello_world.xlsx")
    local worksheet = workbook:add_worksheet()

    worksheet:write("A1", "Hello world")

    workbook:close()

Save this to a file called ``hello.lua`` and run it as follows::

    $ lua hello.lua

This will output a file called ``hello.xlsx`` which should look something like
the following:

.. image:: _images/hello01.png

If you downloaded a tarball or cloned the repo, as shown above, you should also
have a directory called
`examples <https://github.com/jmcnamara/xlsxwriter.lua/tree/master/examples>`_
with some sample applications that demonstrate different features of
xlsxwriter.


Documentation
-------------

The latest version of this document is hosted on
`Read The Docs <http://xlsxwriterlua.readthedocs.org>`_. It is also
available as a
`PDF <http://github.com/jmcnamara/xlsxwriter.lua/blob/master/docs/xlsxwriter_lua.pdf?raw=true>`_.

Once you are happy that the module is installed and operational you can have a
look at the rest of the ``xlsxwriter`` documentation. :ref:`tutorial1` is a good
place to start.
