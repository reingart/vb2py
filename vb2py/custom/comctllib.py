"""Classes to mimic the ComctlLib library from Microsoft"""


# << Constants >>
tvwChild = 4
tvwFirst = 0
tvwLast = 1
tvwManual = 1 
tvwNext = 2
tvwPictureText = 1
tvwPlusMinusText = 2 
tvwPlusPictureText = 3 
tvwPrevious = 3
tvwRootLines = 1
tvwTextOnly = 0
tvwTreeLines = 0
tvwTreelinesPictureText = 5
tvwTreelinesPlusMinusPictureText = 7
tvwTreelinesPlusMinusText = 6
tvwTreelinesText = 4
# -- end -- << Constants >>
# << Classes >>
class Node(object): 
    """A node in a tree view""" 

    # << class Node declarations >>
    pass
    # -- end -- << class Node declarations >> 
    # << class Node methods >> (1 of 3)
    def __init__(self, id=None, parent=None): 
        """Initialise the Node instance"""
        self._id = id
        self._parent = parent
    # << class Node methods >> (2 of 3)
    def _getKey(self): 
        """Get the key"""
        return self._parent.GetPyData(self._id)

    def _setKey(self, key): 
        """Set the key"""
        self._parent.SetPyData(self._id, key)

    Key = property(fget=_getKey,
                   fset=_setKey)
    # << class Node methods >> (3 of 3)
    def _getExpanded(self): 
        """Get the expanded state"""
        return self._parent.IsExpanded(self._id)

    def _setExpanded(self, expanded):
        """Set the expanded state"""
        if expanded:
            self._parent.Expand(self._id)
        else:
            self._parent.Collapse(self._id)

    Expanded = property(fget=_getExpanded,
                        fset=_setExpanded)
    # -- end -- << class Node methods >>
# -- end -- << Classes >> 

if __name__ == "__main__": 
    pass
