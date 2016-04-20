from Constants import *

from Calculator15 import ORSCalculator
import CalculatorAux

class pos(object):
    def __init__(self, **kwargs):
        self.x = kwargs.get("x", kwargs.get("col", 0))
        self.y = kwargs.get("y", kwargs.get("row", 0))
        self.row = self.y
        self.col = self.x
        
# Include my modules
from Parser import *
from Output import *
