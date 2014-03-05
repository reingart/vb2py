from threading import Thread
from time import sleep

def change1(x, y):
    """Mimic ByRef by rebinding after the fact"""
    x = x + 1
    y = y + 1
    return y

def change2(x, y, refs):
    """Mimic ByRef by manipulating locals"""
    x = x + 1
    refs[y] = refs[y] + 1


if __name__ == "__main__":

    def tryit():
        x = 1
        y = 1
        def doit(n, locs, delay):
            print "Thread %d before x, y = %s, %s" % (n, x, y)
            change2(x, 'y', locs)
            print "Thread %d sleeping for %s" % (n, delay)
            sleep(delay)
            print "Thread %d after x, y = %s, %s" % (n, x, y)

        for delay in (0, .00001, 1):	
            t1 = Thread(target=doit, args=(1, locals(), delay))
            t2 = Thread(target=doit, args=(2, locals(), delay))

            t1.start()
            t2.start()

    tryit()
