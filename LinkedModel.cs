using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using _00_RevitLibrary;

namespace SOM.RevitTools.PlaceDoors
{
    class LinkedModel
    {
        /// <summary>
        /// Get linked model and select all the doors. 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="uiApp"></param>
        /// <returns>Object of doors</returns>
        public List<ObjDoors> LinkedModelDoors(Document doc, UIApplication uiApp)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            IList<Element> elems = collector
                .OfCategory(BuiltInCategory.OST_RvtLinks)
                .OfClass(typeof(RevitLinkType))
                .ToElements();

            List<ObjDoors> LinkedModelDoors = new List<ObjDoors>();
            foreach (Element e in elems)
            {
                RevitLinkType linkType = e as RevitLinkType;
                String s = String.Empty;

                foreach (Document linkedDoc in uiApp.Application.Documents)
                {
                    if (linkedDoc.Title.Equals(linkType.Name))
                    {
                        MyLibrary library = new MyLibrary();
                        List<Element> doors = library.GetFamilyElement(linkedDoc, BuiltInCategory.OST_Doors);
                        foreach (Element door in doors)
                        {
                            try
                            {
                                LocationPoint location = door.Location as LocationPoint;
                                // Get level id parameter
                                Parameter levelId = door.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM);
                                // Get level value of id parameter
                                string parmLevel_Id = library.GetParameterValue(levelId);
                                // get all levels in linked model 
                                List<Level> levels = library.GetLevels(linkedDoc);
                                // find and match of door level and linked model level to get level 
                                Level levelName = levels.Find(x => Convert.ToString(x.Id.IntegerValue) == parmLevel_Id);

                                FamilySymbol familySymbol = library.GetFamilySymbol(linkedDoc, door.Name, BuiltInCategory.OST_Doors);
                                // find unique identification of door

                                ObjDoors ObjectDoor = new ObjDoors();
                                ObjectDoor.doorElement = door;
                                ObjectDoor.doorName = door.Name;
                                ObjectDoor.doorId = door.Id.ToString();
                                ObjectDoor.doorWidth = door.get_Parameter(BuiltInParameter.DOOR_WIDTH).AsDouble();
                                ObjectDoor.doorHeight = door.get_Parameter(BuiltInParameter.DOOR_HEIGHT).AsDouble();
                                ObjectDoor.X = location.Point.X;
                                ObjectDoor.Y = location.Point.Y;
                                ObjectDoor.Z = location.Point.Z;
                                ObjectDoor.level = levelName;
                                LinkedModelDoors.Add(ObjectDoor);
                            }
                            catch (Exception er)
                            {

                            }
                        }
                    }
                }
            }
            return LinkedModelDoors;
        }
    }
}
