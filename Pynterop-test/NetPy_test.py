import sys
path = u'K:\\ObjectARX 2013\\inc\\'
sys.path.append(path)


import clr
from CLR.System.Reflection import Assembly
Assembly.LoadFrom(path + 'AcDbMgd.dll')
Assembly.LoadFrom(path + 'AcCoreMgd.dll')
Assembly.LoadFrom(path + 'AcMgd.dll')
#Assembly.LoadWithPartialName('Autodesk.AutoCAD.DatabaseServices')
#clr.AddReference('Autodesk.AutoCAD.Runtime')
#clr.AddReference(path + 'accoremgd.dll')


#ie = Autodesk.AutoCAD.Runtime.IExtensionApplication
#print('Hello AutoCAD')
#import sys
#sys.stdout.write('\a')
#sys.stdout.flush()

#import win32com.client
#acad = win32com.client.Dispatch("AutoCAD.Application")
#help(acad)

import Autodesk.AutoCAD.DatabaseServices as ads
import Autodesk
import Autodesk.AutoCAD.Runtime as ar
import Autodesk.AutoCAD.ApplicationServices as aas
import Autodesk.AutoCAD.DatabaseServices as ads
import Autodesk.AutoCAD.Geometry as ag
import Autodesk.AutoCAD.Internal as ai
from Autodesk.AutoCAD.Internal import Utils

doc = Application.DocumentManager.MdiActiveDocument
ed = doc.Editor
ed.WriteMessage("\nRegistered Python command")

def autocad_command(function):
 
    # First query the function name
    n = function.__name__
    
    # Create the callback and add the command
    cc = ai.CommandCallback(function)
    Utils.AddCommand('pycmds', n, n, ar.CommandFlags.Modal, cc)
 
    # Let's now write a message to the command-line
    #doc = aas.Application.DocumentManager.MdiActiveDocument
    #ed = doc.Editor
    #ed.WriteMessage("\nRegistered Python command: {0}", n)
 
# A simple "Hello World!" command
 
@autocad_command
def msg():
    
    doc = aas.Application.DocumentManager.MdiActiveDocument
    ed = doc.Editor
    ed.WriteMessage("\nOur test command works!")
 
# And one to do something a little more complex...
# Adds a circle to the current space
 
@autocad_command
def mycir():
 
    doc = aas.Application.DocumentManager.MdiActiveDocument
    db = doc.Database
 
    tr = doc.TransactionManager.StartTransaction()
    bt = tr.GetObject(db.BlockTableId, ads.OpenMode.ForRead)
    btr = tr.GetObject(db.CurrentSpaceId, ads.OpenMode.ForWrite)
 
    cir = ads.Circle(ag.Point3d(10,10,0),ag.Vector3d.ZAxis, 2)
 
    btr.AppendEntity(cir)
    tr.AddNewlyCreatedDBObject(cir, True)

    cir = ads.Circle(ag.Point3d(10,10,0),ag.Vector3d.ZAxis, 3)
 
    btr.AppendEntity(cir)
    tr.AddNewlyCreatedDBObject(cir, True)

 
    tr.Commit()
    tr.Dispose()