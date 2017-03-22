using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using DBLibrary;
//using _00_RevitLibrary;

namespace SOM.RevitTools.PlaceDoors
{
    class LocalModel
    {
        public List<ObjDoors> CurrentModelDoors(Document doc)
        {

            List<ObjDoors> currentModelDoors = new List<ObjDoors>();

            {
                LibraryGetItems library = new LibraryGetItems();
                List<Element> doors = library.GetFamilyElement(doc, BuiltInCategory.OST_Doors);
                foreach (Element door in doors)
                {
                    LocationPoint location = door.Location as LocationPoint;
                    // Get level id parameter
                    Parameter levelId = door.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM);
                    // Get level value of id parameter
                    string parmLevel_Id = library.GetParameterValue(levelId);
                    // get all levels in current model 
                    List<Level> levels = library.GetLevels(doc);
                    // find and match of door level and current model level to get level
                    Level levelName = levels.Find(x => Convert.ToString(x.Id.IntegerValue) == parmLevel_Id);

                    FamilySymbol familySymbol = library.GetFamilySymbol(doc,  door.Name, BuiltInCategory.OST_Doors);
                    // find unique identification of door
                    //1bb68248-4b34-4821-9be3-8ddf47852209 = SOM ID
                    var SOMIDParam = door.LookupParameter("SOM ID");
                    string doorId = library.GetParameterValue(SOMIDParam);

                    ObjDoors ObjectDoor = new ObjDoors();
                    ObjectDoor.doorElement = door;
                    ObjectDoor.doorName = door.Name;
                    ObjectDoor.doorId = doorId;
                    ObjectDoor.doorWidth = familySymbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM).AsDouble();
                    ObjectDoor.doorHeight = familySymbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM).AsDouble();
                    ObjectDoor.X = location.Point.X;
                    ObjectDoor.Y = location.Point.Y;
                    ObjectDoor.Z = location.Point.Z;
                    ObjectDoor.level = levelName;
                    currentModelDoors.Add(ObjectDoor);
                }
            }
            return currentModelDoors;
        }
    }
}
