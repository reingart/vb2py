# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

from unittest import *
from vb2py.vbfunctions import vbObjectInitialize, String, Integer

class TestObjectDef(TestCase):

	# << ObjectDef tests >> (1 of 8)
	def test1D(self):
		"""test1D: should be able to create a 1D Array"""
		a = vbObjectInitialize((10,), String)		
		self.assertEqual(len(a), 11)
		for i in range(0, 11):
			self.assertEqual(a(i), "")
	# << ObjectDef tests >> (2 of 8)
	def test1DOffsetRange(self):
		"""test1DOffsetRange: should be able to create a 1D Array with an offset range"""
		a = vbObjectInitialize(((5,10),), String)		
		self.assertEqual(len(a), 6)
		for i in range(5, 11):
			self.assertEqual(a(i), "")
		self.assertRaises(IndexError, a.__getitem__, 4)	
		self.assertRaises(IndexError, a.__getitem__, 11)
	# << ObjectDef tests >> (3 of 8)
	def test2D(self):
		"""test2D: should be able to create a 2D Array"""
		a = vbObjectInitialize((10, 5), String)		
		self.assertEqual(len(a), 11)
		self.assertEqual(len(a(0)), 6)
		for i in range(0, 11):
			for j in range(0, 6):
				self.assertEqual(a(i, j), "")
	# << ObjectDef tests >> (4 of 8)
	def test2DOffsetRange(self):
		"""test2DOffsetRange: should be able to create a 2D Array with offset range"""
		a = vbObjectInitialize(((5, 10), (-5, 5)), String)		
		self.assertEqual(len(a), 6)
		self.assertEqual(len(a(5)), 11)
		for i in range(5, 11):
			for j in range(-5, 6):
				self.assertEqual(a(i, j), "")
	# << ObjectDef tests >> (5 of 8)
	def test1DIteration(self):
		"""test1DIteration: should be able to iterate over a 1D Array"""
		a = vbObjectInitialize((10,), Integer)		
		for i in range(0, 11):
			a[i] = i
		for i, j in zip(range(11), a):
			self.assertEqual(i, j)
	# << ObjectDef tests >> (6 of 8)
	def test1DIterationOffsetRange(self):
		"""test1DIterationOffsetRange: should be able to iterate over a 1D Array with offset range"""
		a = vbObjectInitialize(((5, 10),), Integer)		
		for i in range(5, 11):
			a[i] = i
		for i, j in zip(range(5, 11), a):
			self.assertEqual(i, j)
	# << ObjectDef tests >> (7 of 8)
	def test2DIteration(self):
		"""test2DIteration: should be able to iterate over a 2D Array"""
		a = vbObjectInitialize((10,20), Integer)		
		for i in range(0, 11):
			for j in range(0, 21):
				a[i, j] = i*j
		for i, arr in zip(range(11), a):
			for j, result in zip(range(21), arr):
				self.assertEqual(i*j, result)
	# << ObjectDef tests >> (8 of 8)
	def test2DIterationOffsetRange(self):
		"""test2DIterationOffsetRange: should be able to iterate over a 2D Array with an offset range"""
		a = vbObjectInitialize(((5, 10),(10, 20)), Integer)		
		for i in range(5, 11):
			for j in range(10, 21):
				a[i, j] = i*j
		for i, arr in zip(range(5, 11), a):
			for j, result in zip(range(10, 21), arr):
				self.assertEqual(i*j, result)
	# -- end -- << ObjectDef tests >>


if __name__ == "__main__":
	main()
