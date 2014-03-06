import os
import sys

# << Utilities >>
def rootPath():
    """Return the root path"""
    return os.path.dirname(__file__)


def relativePath(path):
    """Return the path to a file"""
    return os.path.join(rootPath(), path)
# -- end -- << Utilities >>
