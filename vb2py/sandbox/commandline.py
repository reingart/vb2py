# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

"""A simple module intended to aid interactive testing of the parser

We just import a lot of useful things with short names so they are easy to type!

"""


from vb2py.vbparser import convertVBtoPython, parseVB as p, parseVBFile as f, getAST as t
import vb2py.vbparser

try:
    from win32clipboard import *
    import win32con

    def getClipBoardText():
        """Get text from the clipboard"""
        OpenClipboard()
        try:
            got = GetClipboardData(win32con.CF_UNICODETEXT)
        finally:
            CloseClipboard()
        return str(got)
    v = getClipBoardText
except ImportError:
    print "Clipboard copy not working!"

if __name__ == "__main__":
    def c(*args, **kw):
        print convertVBtoPython(*args, **kw)
