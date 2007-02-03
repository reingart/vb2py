"""Install vb2Py

Much of the following code is copied from the PythonCard installation script
because the original setup.py would copy files to all sorts of weird location
on Linux.

You must run this to create the setup distribution from the site-packages!

"""


WIN_DEFAULT_COMMAND = "install"
APPLICATION_NAME = "vb2py"
from distutils.core import setup
from distutils.command.install_data import install_data
import glob, os, sys
if len(sys.argv) == 1 and sys.platform.startswith("win"):
    sys.argv.append(WIN_DEFAULT_COMMAND)

"""
This script is setup.py of the vb2py package.

"""

# << Support functions >>
class smart_install_data(install_data):
    def run(self):
        #need to change self.install_dir to the actual library dir
        install_cmd = self.get_finalized_command('install')
        self.install_dir = getattr(install_cmd, 'install_lib')
        return install_data.run(self) 

def recurseDir(startDir):
    # This should all be replaced by calls to os.path.walk, but later
    listX=[startDir]
    for fyle in os.listdir(startDir):
        file=os.path.join(startDir,fyle)
        if os.path.isdir(file):
            listX.extend(recurseDir(file))
    return listX

def makeDataDirs(rootDir=APPLICATION_NAME, dataDirs=[]):
    "Construct a list of the data directories to be included"
    # This function will return a list of tuples, each tuple being of the form;
    #  ( <target_directory_name>, [<list_of_files>] )
    listX=[]
    results=[]
    for directory in dataDirs:
        directories=recurseDir(directory)
        results.extend(directories)
    for directory in results:
        if os.path.split(directory)[1]!='CVS':
            # Add this directory and its contents to list
            files=[]
            for file in os.listdir(directory):
                if file!='CVS' and file!='.cvsignore' and os.path.splitext(file)[1].lower() <> ".htm":
                    if os.path.isfile(os.path.join(directory, file)):
                        files.append(os.path.join(directory, file))
            listX.append((rootDir+'/'+directory, files))
    # list.append((rootDir, 'stc_styles.cfg'))

    return listX
# -- end -- << Support functions >>

setup(name=APPLICATION_NAME, 
      version="0.2.1",
      description="Visual Basic to Python Converter",
      author="Paul Paterson",
      author_email="paulpaterson@users.sourceforge.net",
      url="http://vb2py.sourceforge.net",
      packages=["vb2py", "vb2py.test", "vb2py.sandbox", 
                "vb2py.plugins", "vb2py.targets", "vb2py.targets.pythoncard"],
      package_dir={APPLICATION_NAME: '.'},
      license="BSD",
      cmdclass = { 'install_data': smart_install_data},
      data_files=makeDataDirs(dataDirs=[".", "./test", "./vb", "./doc", "./sandbox", "./targets", "./vb"]),
     )
