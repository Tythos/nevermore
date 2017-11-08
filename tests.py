"""
"""

import unittest
import nevermore

class TestModel(object):
    """
    """
    
    def __init__(self):
        """
        """
        self.name = 'Brian'
        self.age = 32
        
class BaseTests(unittest.TestCase):
    """
    """
    
    def test_openmetaclose(self):
        """
        """
        with nevermore.DataStore("test.xlsx") as ds:
            ds._addTable(nevermore.Meta)
        
    def test_create(self):
        """
        """
        with nevermore.DataStore("test.xlsx") as ds:
            ds.create(TestModel())
            
if __name__ == "__main__":
    unittest.main()
