#region Namespaces
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

#endregion

namespace SOM.RevitTools.PlaceDoors
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            // Revit application documents. 
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
 
            program Program = new program();
            // Linked model doors 
            List<ObjDoors> List_DoorsLinkedModel = new List<ObjDoors>();
            LinkedModel linkedModel = new LinkedModel();
            
            // Current model doors
            List<ObjDoors> List_DoorsCurrentModel = new List<ObjDoors>();
            LocalModel localModel = new LocalModel();

            // get doors from linked model 
            List_DoorsLinkedModel = linkedModel.LinkedModelDoors(doc, uiapp);
            ExportToExcel(List_DoorsLinkedModel);
            // get doors from current model. 
            List_DoorsCurrentModel = localModel.CurrentModelDoors(doc);

            Program.DoorProgram(doc, uidoc, List_DoorsLinkedModel, List_DoorsCurrentModel);
            return Result.Succeeded;
        }
    }
}
