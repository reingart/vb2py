The main documentation lives at .\doc\index.html


vb2Py Installation
==================

**Important** If you installed v0.1, or the CVS version, prior to v0.2 please remove the old directories completely before installing v0.2. Changes in the package (in particular the renaming of ``vb2py.py``) **will** causes problems. Sorry for the confusion!

* `Main Installation`_
* `GUI Installation`_
* `Simpleparse Installation`_
* `mxTools Installation`_
* `PythonCard Installation`_

``vb2Py`` uses Python and has been tested on Python 2.2. Python 2.3 should work but earlier versions will not.
 

Main Installation
~~~~~~~~~~~~~~~~~

vb2Py is written in Python and runs on any platform which has a Python interpreter.

Once you have downloaded the ``vb2py`` package you will have a zip file. Before you can do anything you must have Python installed. After Python is installed you can install the ``vb2py`` modules by going to the directory you unzipped the files to and typing::

	> python setup.py install

Now you should have a 'vb2py' folder in your Python site packages directory.

You also need to make sure you have both PythonCard_ and Simpleparse_ installed on your system.

Once these additional resources are installed on your system you should be ready to go and use the converter.

Note: You do not need VB to run the converter!


GUI Installation
~~~~~~~~~~~~~~~~

The vb2Py GUI is also written in Python and uses the PythonCard GUI toolkit. Installation of the GUI is the same as for the main libraries. Once you have downloaded the ``vb2pygui`` you will have a zip file. You can install the ``vb2pygui`` module by going to the directory you unzipped the files to and typing::

	> python setup.py install

Now you should have a 'vb2pygui' folder in your Python site packages directory.


Simpleparse Installation
~~~~~~~~~~~~~~~~~~~~~~~~

vb2Py uses the ``Simpleparse`` module to parse the Visual Basic code. You can download the ``Simpleparse`` files from the Simpleparse_ download site. *NB You need ``Simpleparse v2.0.1a2`` or later.* Follow the instructions on the `Simpleparse homepage`_ to install the software.

You will also need to do the `mxTools Installation`_.


mxTextTools Installation
~~~~~~~~~~~~~~~~~~~~

The ``Simpleparse`` library uses the mxTextTools_ libraries. Once you have downloaded these, follow the instructions on the `mxTools homepage`_.


PythonCard Installation
~~~~~~~~~~~~~~~~~~~~~~~

If you want to use the ``vb2Py`` GUI or view converted forms actually running, then you will need the ``PythonCard`` GUI library. You can download the software from the PythonCard_ download site. Installation instructions are on the `PythonCard homepage`_. ``vb2Py`` has been tested with the 0.7 Prototype version but should work with later versions also.

.. _Simpleparse: http://sourceforge.net/project/showfiles.php?group_id=55673
.. _`Simpleparse homepage`: http://simpleparse.sourceforge.net
.. _mxTextTools: http://www.egenix.com
.. _`mxTools homepage`: http://www.egenix.com
.. _PythonCard: http://sourceforge.net/project/showfiles.php?group_id=19015
.. _`PythonCard homepage`: http://pythoncard.sourceforge.net/installation.html


Command Line Usage
==================

The easiest way to use vb2Py from the command line is by working in the vb2py directory in your Python site-packages location.

To check the basic usage from the command line::

	> python converter.py -h

Two example projects should be installed in the folder, ``vb2py/vb/test1`` and ``test2``. To convert the ``test2`` project (which is slightly more meaningful than ``test1``!) to a PythonCard project, use the command::

	> python converter.py -c vb/test/test2.vbp vb/test/test2
	
	-c = convert the code also	
	vb/test/test2.vbp = the VB project file for the project	
	vb/test/test2 = the location of the generated Python files

The directory where the Python files will be put **must** exist before you run the conversion.

If you have PythonCard_ installed you can see what the forms look like by typing::

	> python vb/test/test2/frmMain.py

You should see a form open with a lot of controls on it. Clicking on the controls should perform the same actions as for the VB version.

.. image:: images/sm_test2_frmmain_py.gif

See the `full form`_

For cases where you don't have a VB project file (``*.vbp``), you can still use the converter to convert individual files on at a time. The following command will convert the single code module in the test2 project::

	> python converter.py -c -f vb/test/Utils.bas vb/test/test2

The code will be placed in the destination directory. Note that when doing file-by-file conversions certain parts of the conversion will not be possible. In particular, global variables will not be correctly handled since in order to determine the correct qualified name for a global, vb2Py needs to know all the files within the project.

For example, when converting an entire project the following code segment (where ``BASEDIR`` is defined in a module called ``FileGlobals``)::

	Function addBaseDirectory(Path as String)
		addBaseDirectory = BASEDIR & Path ' BASEDIR is a global
	End Function

Would be converted to::

	def addBaseDirectory(Path):
		_ret = FileGlobals.BASEDIR + Path
		return _ret

However, if the file is translated on its own (using the ``-f`` command line switch) then it would be converted to::



