#@+leo
#@+node:0::@file sandbox/filetesting.py
#@+body

from __future__ import generators

"""Some testing to see if we can clean up the file access"""


class Test(object):
	
	def __init__(self, dataset):
		self.data = dataset
		self._datastream = iter(self.data)
		self.input = self.next
		
	def __getattribute__(self, name):
		print "Get", name
		return super(Test, self).__getattribute__(name)
	
	def __iter__(self):
		print "Getting iterator"
		return self

	def next(self):
		print "In next"
		d = self._datastream[:]
		while d:
			print "Yielding", d[0]
			yield d.pop()

	def __getitem__(self, index):
		print "Getting item", index
		return self._datastream[index]
	
if __name__ == "__main__":
	f = Test(range(20))
	a, b, c = f
	

#@-body
#@-node:0::@file sandbox/filetesting.py
#@-leo
