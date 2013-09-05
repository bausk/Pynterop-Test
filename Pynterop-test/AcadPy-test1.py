from math import *
import array
import operator
import collections
import pickle
import win32com
import types

from win32com.client import VARIANT
from pythoncom import VT_VARIANT
import pythoncom

def POINT(x,y,z):
   return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
(x,y,z))

def point(*args):
    lst = [0.]*3
    if len(args) < 3:
        lst[0:2] = [float(x) for x in args[0:2]]
    else:
        lst = [float(x) for x in args[0:3]]
    return VARIANT(VT_VARIANT, array.array("d",lst))

class APoint(array.array):
    """ 3D point with basic geometric operations and support for passing as a
        parameter for `AutoCAD` Automation functions

    Usage::

        >>> p1 = APoint(10, 10)
        >>> p2 = APoint(20, 20)
        >>> p1 + p2
        APoint(30.00, 30.00, 0.00)

    Also it supports iterable as parameter::

        >>> APoint([10, 20, 30])
        APoint(10.00, 20.00, 30.00)
        >>> APoint(range(3))
        APoint(0.00, 1.00, 2.00)

    Supported math operations: `+`, `-`, `*`, `/`, `+=`, `-=`, `*=`, `/=`::

        >>> p = APoint(10, 10)
        >>> p + p
        APoint(20.00, 20.00, 0.00)
        >>> p + 10
        APoint(20.00, 20.00, 10.00)
        >>> p * 2
        APoint(20.00, 20.00, 0.00)
        >>> p -= 1
        >>> p
        APoint(9.00, 9.00, -1.00)

    It can be converted to `tuple` or `list`::

        >>> tuple(APoint(1, 1, 1))
        (1.0, 1.0, 1.0)

    """
    def __new__(cls, x_or_seq, y=0.0, z=0.0):
        if isinstance(x_or_seq, (array.array, list, tuple)) and len(x_or_seq) == 3:
            return super(APoint, cls).__new__(cls, 'd', x_or_seq)
        return super(APoint, cls).__new__(cls, 'd', (x_or_seq, y, z))


    @property
    def x(self):
        """ x coordinate of 3D point"""
        return self[0]

    @x.setter
    def x(self, value):
        self[0] = value

    @property
    def y(self):
        """ y coordinate of 3D point"""
        return self[1]

    @y.setter
    def y(self, value):
        self[1] = value

    @property
    def z(self):
        """ z coordinate of 3D point"""
        return self[2]

    @z.setter
    def z(self, value):
        self[2] = value

    def distance_to(self, other):
        """ Returns distance to `other` point

        :param other: :class:`APoint` instance or any sequence of 3 coordinates
        """
        return distance(self, other)

    def __add__(self, other):
        return self.__left_op(self, other, operator.add)

    def __sub__(self, other):
        return self.__left_op(self, other, operator.sub)

    def __mul__(self, other):
        return self.__left_op(self, other, operator.mul)

    def __div__(self, other):
        return self.__left_op(self, other, operator.div)

    __radd__ = __add__
    __rsub__ = __sub__
    __rmul__ = __mul__
    __rdiv__ = __div__

    def __neg__(self):
        return self.__left_op(self, -1, operator.mul)

    def __left_op(self, p1, p2, op):
        if isinstance(p2, (float, int)):
            return APoint(op(p1[0], p2), op(p1[1], p2), op(p1[2], p2))
        return APoint(op(p1[0], p2[0]), op(p1[1], p2[1]), op(p1[2], p2[2]))

    def __iadd__(self, p2):
        return self.__iop(p2, operator.add)

    def __isub__(self, p2):
        return self.__iop(p2, operator.sub)

    def __imul__(self, p2):
        return self.__iop(p2, operator.mul)

    def __idiv__(self, p2):
        return self.__iop(p2, operator.div)

    def __iop(self, p2, op):
        if isinstance(p2, (float, int)):
            self[0] = op(self[0], p2)
            self[1] = op(self[1], p2)
            self[2] = op(self[2], p2)
        else:
            self[0] = op(self[0], p2[0])
            self[1] = op(self[1], p2[1])
            self[2] = op(self[2], p2[2])
        return self

    def __repr__(self):
        return self.__str__()

    def __str__(self):
        return 'APoint(%.2f, %.2f, %.2f)' % tuple(self)

    def __eq__(self, other):
        if not isinstance(other, (array.array, list, tuple)):
            return False
        return tuple(self) == tuple(other)


def variant(data):
    return VARIANT(VT_VARIANT, data)

def vararr(*data):
    if (  len(data) == 1 and 
          isinstance(data, collections.Iterable) ):
        data = data[0]
    return map(variant, data)

def _sequence_to_comtypes(typecode='d', *sequence):
    if len(sequence) == 1:
        return array.array(typecode, sequence[0])
    return array.array(typecode, sequence)

def main():
    test2()


def test1():
    #Test1: mesh testing

    acad = Autocad()
    acad.prompt("Hello, Autocad from Python\n")
    print acad.doc.Name

    p1 = APoint(0, 0)
    p2 = APoint(50, 25)
    for i in range(5):
        text = acad.model.AddText('Hi %s!' % i, p1, 2.5)
        acad.model.AddLine(p1, p2)
        acad.model.AddCircle(p1, 10)
        p1.y += 10

    dp = APoint(10, 0)
    for text in acad.iter_objects('Text'):
        print('text: %s at: %s' % (text.TextString, text.InsertionPoint))
        text.InsertionPoint = APoint(text.InsertionPoint) + dp

    for obj in acad.iter_objects(['Circle', 'Line']):
        print(obj.ObjectName)

def test2():
    #Test2: Direct AutoCAD COM
    p3 = POINT(20.00, 20.00, 0.00)
    p2 = APoint(10,10,0)
    appObj = win32com.client.Dispatch("AutoCAD.Application")  
    docObj = appObj.ActiveDocument
    modelSpaceObj = docObj.ModelSpace

    p1 = _sequence_to_comtypes('d', 10, 10)
    for i in range(5):
        p1 = point(10,10)
        modelSpaceObj.AddCircle(p3, 10)
        #p1.y += 10

    #dp = APoint(10, 0)
    #for text in acad.iter_objects('Text'):
    #    print('text: %s at: %s' % (text.TextString, text.InsertionPoint))
    #    text.InsertionPoint = APoint(text.InsertionPoint) + dp
    #
    #for obj in acad.iter_objects(['Circle', 'Line']):
    #    print(obj.ObjectName)

if __name__ == '__main__':
    main()