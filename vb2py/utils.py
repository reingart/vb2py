import os
import sys

# << Utilities >>
def rootPath():
    """Return the root path"""
    return os.path.join(os.path.abspath(__file__).split("vb2py")[0], "vb2py")


def relativePath(path):
    """Return the path to a file"""
    return os.path.join(rootPath(), path)
# -- end -- << Utilities >>
