# Created by Leo from: C:\Development\Python23\Lib\site-packages\vb2py\vb2py.leo

import glob
import os
import unittest
import sys
import re
import time

ok = re.compile(r".*Ran\s(\d+).*", re.DOTALL+re.MULTILINE)
fail = re.compile(r".*Ran\s(\d+).*failures=(\d+)", re.DOTALL+re.MULTILINE)

show_errors = 0
if len(sys.argv) == 2:
	if sys.argv[1] == "-v":
		show_errors = 1

if __name__ == "__main__":
	print "\nStarting testall at %s\n" % time.ctime()
	files = glob.glob(r"test/test*.py")
	total_run = 0
	total_failed = 0
	start = time.time()
	try:
		for file in files:
			if file not in (r"test\testall.py", r"test\testframework.py", "test\testparser.py"):
				print "Running '%s' ... " % file,
				fname = os.path.join(r"c:\development\python23\lib\site-packages\vb2py", file)
				pi, po, pe = os.popen3(fname)
				result = pe.read()
				if result.find("FAILED") > -1:
					try:
						num = int(fail.match(result).groups()[0])
						num_fail = int(fail.match(result).groups()[1])
					except:
						num, num_fail = 0, 0
					total_run += num
					total_failed += num_fail
					if show_errors:
									print "\n%s" % result
					else:
						print "*** %s errors out of %s" % (num_fail, num)
				else:
					num = int(ok.match(result).groups()[0])
					print "Passed %s tests" % num
					total_run += num
				pi.close()
				po.close()
				pe.close()
	except KeyboardInterrupt:
		pass
	print "\nRan %d tests\nFailed %d\nTook %d seconds" % (total_run, total_failed, time.time()-start)
