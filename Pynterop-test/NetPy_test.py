import sys
path = u'K:\\ObjectARX 2013\\inc\\'
sys.path.append(path)


import clr
from CLR.System.Reflection import Assembly
Assembly.LoadFrom(path + 'AcDbMgd.dll')
Assembly.LoadFrom(path + 'AcCoreMgd.dll')
Assembly.LoadFrom(path + 'AcMgd.dll')

import Autodesk.AutoCAD.DatabaseServices as ads
import Autodesk
import Autodesk.AutoCAD.Runtime as ar
import Autodesk.AutoCAD.ApplicationServices as aas
import Autodesk.AutoCAD.DatabaseServices as ads
import Autodesk.AutoCAD.Geometry as ag
import Autodesk.AutoCAD.Internal as ai
from Autodesk.AutoCAD.Internal import Utils

def msg():
    
    doc = aas.Application.DocumentManager.MdiActiveDocument
    ed = doc.Editor
    ed.WriteMessage("\nOur test command works!")

cc = ai.CommandCallback(msg)
#n = "msg"
#Utils.AddCommand('pycmds', n, n, ar.CommandFlags.Modal, cc)
doc = aas.Application.DocumentManager.MdiActiveDocument
ed = doc.Editor
ed.WriteMessage("\nRegistered Python command")

