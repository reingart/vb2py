from testframework import *

# << Operators tests >> (1 of 2)
tests.extend([
    ('a = "the dog" Like "*dog"', {"a" : 1}),
    ('a = "the big dog" Like "*dog"', {"a" : 1}),
    ('a = "the big dooog" Like "*dog"', {"a" : 0}),
    ('a = "the big dooog" Like "the*"', {"a" : 1}),
    ('a = "   the big dooog" Like "the*"', {"a" : 0}),
    ('a = "   the big dooog" Like "*the*"', {"a" : 1}),
    ('a = "the doggy" Like "*dog??"', {"a" : 1}),
    ('a = "the doggy" Like "*dog???"', {"a" : 0}),
    ('a = "the doggy" Like "*dog?"', {"a" : 0}),
    ('a = "the dogs" Like "*dog?"', {"a" : 1}),
    ('a = "the dogs" Like "??? dog?"', {"a" : 1}),
    ('a = "them dogs" Like "??? dog?"', {"a" : 0}),

    ('a = "the" & "dog" Like "???dog"', {"a" : 1}),
    ('a = "one" & "the" & "dog" Like "???dog"', {"a" : 0}),

])
# << Operators tests >> (2 of 2)
tests.extend([
    ('a = 0 Xor 0', {"a" : 0}),
    ('a = 1 Xor 0', {"a" : 1}),
    ('a = 0 Xor 1', {"a" : 1}),
    ('a = 1 Xor 1', {"a" : 0}),

    ('a = 10 Xor 2', {"a" : 8}),
    ('a = 100 Xor 254', {"a" : 154}),

])
# -- end -- << Operators tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
    main()
