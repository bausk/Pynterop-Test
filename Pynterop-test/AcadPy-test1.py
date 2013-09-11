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