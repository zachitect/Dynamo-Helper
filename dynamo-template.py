#Contact: Zach X.G. Zheng
#Email: Zach.Zheng@Jacobs.com

# Enable DotNet via Common Language Runtime
import clr
import math

# Import RevitAPI
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
# from Autodesk.Revit.DB.Structure import *

# Import RevitAPIUI
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import *

# Import System
clr.AddReference('System')
from System.Collections.Generic import List

# Import Revit Nodes
clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.GeometryConversion)
clr.ImportExtensions(Revit.Elements)

# Import DesignScript
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

# Import Revit Services & Transaction
clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

# Pointing the current Document
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = uidoc.ActiveView
