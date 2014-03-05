from vb2py.vbclasses import Collection
import unittest 

class TestCollection(unittest.TestCase):

    def setUp(self):
        """Set up the test"""
        self.c = Collection()

    # << Collection tests >> (1 of 9)
    def testAddNumeric(self):
        """testAddNumeric: should be able to add with numeric indexes"""
        for i in range(10):
            self.c.Add(i)
        for expect, actual in zip(range(10), self.c):
            self.assertEqual(expect, actual)
        self.assertEqual(self.c.Count(), 10)
    # << Collection tests >> (2 of 9)
    def testAddBeforeNumeric(self):
        """testAddBeforeNumeric: should be able to add something before something else"""
        # Put 1 ... 9 in with 5 missing
        for i in range(1, 10):
            if i <> 5:
                self.c.Add(i)
        self.c.Add(5, Before=5) # ie before the index 5
        for expect, actual in zip(range(1, 10), self.c):
            self.assertEqual(expect, actual)
        self.assertEqual(self.c.Count(), 9)
    # << Collection tests >> (3 of 9)
    def testAddAfterNumeric(self):
        """testAddAfterNumeric: should be able to add something after something else"""
        # Put 1 ... 9 in with 5 missing
        for i in range(1, 10):
            if i <> 5:
                self.c.Add(i)
        self.c.Add(5, After=4)
        for expect, actual in zip(range(1, 10), self.c):
            self.assertEqual(expect, actual)
        self.assertEqual(self.c.Count(), 9)
    # << Collection tests >> (4 of 9)
    def testAddText(self):
        """testAddText: should be able to add with text indexes"""
        for i in range(10):
            self.c.Add(i, "txt%d" % i)
        for expect, actual in zip(range(10), self.c):
            self.assertEqual(expect, actual)
        self.assertEqual(self.c.Count(), 10)
    # << Collection tests >> (5 of 9)
    def testAddTextandNumeric(self):
        """testAddTextandNumeric: should be able to add with text and numeric indexes"""
        for i in range(10):
            self.c.Add(i, "txt%d" % i)
            self.c.Add(i)
        for i in range(10):
            self.assertEqual(self.c.Item("txt%d" % i), i)
            self.assertEqual(self.c.Item(i*2+2), i)
        self.assertEqual(self.c.Count(), 20)
    # << Collection tests >> (6 of 9)
    def testItemNumeric(self):
        """testItemNumeric: should be able to get with numeric indexes"""
        for i in range(10):
            self.c.Add(i)
        for i in range(10):
            self.assertEqual(i, self.c.Item(i+1))
    # << Collection tests >> (7 of 9)
    def testItemText(self):
        """testItemText: should be able to get with text indexes"""
        for i in range(10):
            self.c.Add(i, "txt%d" % i)
        for i in range(10):
            self.assertEqual(i, self.c.Item("txt%d" %  i))
    # << Collection tests >> (8 of 9)
    def testRemoveNumeric(self):
        """testRemoveNumeric: should be able to remove with numeric indexes"""
        for i in range(10):
            self.c.Add(i+1)
        self.c.Remove(5)
        self.assertEqual(self.c.Count(), 9)
        for i in self.c:
            self.assertNotEqual(i, 5)
    # << Collection tests >> (9 of 9)
    def testRemoveText(self):
        """testRemoveText: should be able to remove with text indexes"""
        for i in range(10):
            self.c.Add(i, "txt%d" % i)
        self.c.Remove("txt%d" % 5)
        self.assertEqual(self.c.Count(), 9)
        for i in self.c:
            self.assertNotEqual(i, 5)
    # -- end -- << Collection tests >>


if __name__ == "__main__":
    unittest.main()
