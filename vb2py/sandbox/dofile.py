from vb2py.vbparser import convertVBtoPython, parseVB as p
import vb2py.vbparser
import sys

if __name__ == "__main__":
    def c(*args, **kw):
        print convertVBtoPython(*args, **kw)

    t = open(sys.argv[1], 'r').read()
