using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.IO;
using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;
using Autodesk.Revit.DB.Structure;
using _00_RevitLibrary;
using NLog;

namespace SOM.RevitTools.PlaceDoors
{
    class program
    {
        //Fields 
        public List<ObjDoors> _List_CreatedDoors {get; set;}

        /// <summary>
        /// NLog added to log errors. 
        /// </summary>
        private static Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Create or move doors.  Checks to see if the door exist. 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="uidoc"></param>
        /// <param name="List_DoorsLinkedModel"></param>
        /// <param name="List_DoorsCurrentModel"></param>
        public void DoorProgram(Document doc, UIDocument uidoc,
            List<ObjDoors> List_DoorsLinkedModel, List<ObjDoors> List_DoorsCurrentModel)
        {
            // logger object. 
            Logger logger = LogManager.GetLogger("program");
            ExportExcel exportExcel = new ExportExcel();

            foreach (ObjDoors linkedDoor in List_DoorsLinkedModel)
            {
                // check to see if door exist 
                ObjDoors DoorFound = List_DoorsCurrentModel.Find(x => x.doorId == linkedDoor.doorId);
                // if it doesn't exist it will create a new door. 
                if (DoorFound == null)
                {
                    try
                    {
                        CreateDoors(uidoc, doc, linkedDoor);
                    }
                    catch (Exception e)
                    {
                        string ErrMessage = e.Message;
                    }
                }

                // if door exist the check to see if the location is the same and type. 
                if (DoorFound != null)
                {
                    try
                    {
                        MoveDoors(doc, linkedDoor, DoorFound);
                    }
                    catch (Exception e)
                    {
                        string ErrMessage = e.Message;
                    }

                    // Check to see if door size match. 
                    double height = Math.Round(linkedDoor.doorHeight, 2);
                    double width = Math.Round(linkedDoor.doorWidth, 2);
                    String doorType = height.ToString() + "ft" + " x " + width.ToString() + "ft";
                    if (doorType != "0ft x 0ft")
                    {
                        if (DoorFound.doorName != doorType)
                        {
                            //FamilySymbol familySymbol = findSymbol(doc, DoorFound, doorType);
                            FamilySymbol familySymbol = FindElementByName(doc, typeof(FamilySymbol), doorType) as FamilySymbol;
                            if (familySymbol == null)
                            {
                                FamilySymbol oldType = findSymbol(doc, DoorFound);

                                FamilySymbol ChangeFamilySymbol = CreateNewType(doc, oldType, linkedDoor);
                                changeType(doc, DoorFound, ChangeFamilySymbol);
                            }
                            if (familySymbol != null)
                            {
                                changeType(doc, DoorFound, familySymbol);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Create a door in the project from linked model.  
        /// </summary>
        /// <param name="uidoc"></param>
        /// <param name="doc"></param>
        /// <param name="linkedDoor"></param>
        public void CreateDoors(UIDocument uidoc, Document doc, ObjDoors linkedDoor)
        {
            MyLibrary Library = new MyLibrary();
            Level level = findLevel(doc, linkedDoor);
            // make door type height x width 
            double height = Math.Round(linkedDoor.doorHeight, 2);
            double width = Math.Round(linkedDoor.doorWidth, 2);
            String doorType = height.ToString() + "ft" + " x " + width.ToString() + "ft";

            if (doorType != "0ft x 0ft")
            {
                FamilySymbol currentDoorType = null;
                try
                {
                    currentDoorType = FindElementByName(doc, typeof(FamilySymbol), doorType) as FamilySymbol;
                }
                catch { }

                if (currentDoorType == null)
                {
                    FamilySymbol familySymbol_OldType = FindElementByName(doc, typeof(FamilySymbol), "12ft x 12ft") as FamilySymbol;
                    currentDoorType = CreateNewType(doc, familySymbol_OldType, linkedDoor);
                }

                // Convert coordinates to double and create XYZ point.
                XYZ xyz = new XYZ(linkedDoor.X, linkedDoor.Y, linkedDoor.level.Elevation + linkedDoor.doorSillHeight);

                //Find the hosting Wall (nearst wall to the insertion point)
                List<Wall> walls = Library.GetWalls(doc);

                Wall wall = null;
                double distance = double.MaxValue;
                foreach (Element e in walls)
                {
                    Wall w = e as Wall;
                    LocationCurve lc = w.Location as LocationCurve;
                    Curve curve = lc.Curve;
                    XYZ z = curve.GetEndPoint(0);
                    if (linkedDoor.level.Elevation <= z.Z + 3 && linkedDoor.level.Elevation >= z.Z - 3)
                    {
                        int i = w.Id.IntegerValue;
                    }
                    //int i = w.Id.IntegerValue;
                    double proximity = (w.Location as LocationCurve).Curve.Distance(xyz);
                    if (proximity < distance)
                    {
                        distance = proximity;
                        wall = w;
                    }
                }

                // Create door.
                using (Transaction t = new Transaction(doc, "Create door"))
                {
                    
                    t.Start();
                    if (!currentDoorType.IsActive)
                    {
                        // Ensure the family symbol is activated.
                        currentDoorType.Activate();
                        doc.Regenerate();
                    }
                    // Create window
                    // unliss you specified a host, Revit will create the family instance as orphabt object.
                    FamilyInstance fm = doc.Create.NewFamilyInstance(xyz, currentDoorType, wall, level, StructuralType.NonStructural);
                    // Set new local door id to match linked model element id. 
                    Parameter SOMIDParam = fm.LookupParameter("SOM ID");
                    SOMIDParam.Set(linkedDoor.doorId);
                    t.Commit();
                    _List_CreatedDoors.Add(linkedDoor);
                }
            }
        }

        /// <summary>
        /// Move door by comparing XYZ position. 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="linkedDoor"></param>
        /// <param name="localDoor"></param>
        public void MoveDoors(Document doc, ObjDoors linkedDoor, ObjDoors localDoor)
        {
            Element linkedElement = linkedDoor.doorElement;
            Element localElement = localDoor.doorElement;
            // get the column current location
            LocationPoint New_InstanceLocation = linkedElement.Location as LocationPoint;
            LocationPoint Old_InstanceLocation = localElement.Location as LocationPoint;

            XYZ oldPlace = Old_InstanceLocation.Point;

            // XY
            double New_x = linkedDoor.X;
            double New_y = linkedDoor.Y;
            //Get Level
            Level New_level = findLevel(doc, linkedDoor);

            XYZ New_xyz = new XYZ(New_x, New_y, New_level.Elevation);

            double Move_X = New_x - oldPlace.X;
            double Move_Y = New_y - oldPlace.Y;
            //Get Level
            Level Old_level = findLevel(doc, linkedDoor);
            double Move_Z = New_level.Elevation - Old_level.Elevation;

            // Move the element to new location.
            XYZ new_xyz = new XYZ(Move_X, Move_Y, Move_Z);

            //Start move using a transaction. 
            Transaction t = new Transaction(doc);
            t.Start("Move Element");

            ElementTransformUtils.MoveElement(doc, localElement.Id, new_xyz);
            //ElementTransformUtils.RotateElement(doc, element.Id, New_Axis, Rotate);

            t.Commit();
        }

        /// <summary>
        /// Create new type by duplicating old type. 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="oldType"></param>
        /// <param name="linkedDoor"></param>
        /// <returns></returns>
        public FamilySymbol CreateNewType(Document doc, FamilySymbol oldType, ObjDoors linkedDoor)
        {
            double height = Math.Round(linkedDoor.doorHeight, 2);
            double width = Math.Round(linkedDoor.doorWidth, 2);
            String doorType = height.ToString() + "ft" + " x " + width.ToString() + "ft";

            FamilySymbol familySymbol = null;
            using (Transaction t = new Transaction(doc, "Duplicate door"))
            {
                t.Start("duplicate");
                familySymbol = oldType.Duplicate(doorType) as FamilySymbol;
                familySymbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM).Set(width);
                familySymbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM).Set(height);
                t.Commit();
            }
            return familySymbol; 
        }

        /// <summary>
        /// Change Type of element in Revit 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="CurrentDoors"></param>
        /// <param name="symbol"></param>
        public void changeType(Document doc, ObjDoors CurrentDoors, FamilySymbol symbol)
        {
            Element e = CurrentDoors.doorElement;
            FamilyInstance familyInstance = e as FamilyInstance;
            // Transaction to change the element type 
            Transaction trans = new Transaction(doc, "Edit Type");
            trans.Start();
            try
            {
                familyInstance.Symbol = symbol;
            }
            catch { }
            trans.Commit();

        }

        /// <summary>
        /// find level of selected door. 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="door"></param>
        /// <returns></returns>
        public Level findLevel(Document doc, ObjDoors door)
        {
            string levelName = door.level.Name;
            // LINQ to find the level by its name.
            Level level = (from lvl in new FilteredElementCollector(doc).
                               OfClass(typeof(Level)).
                               Cast<Level>()
                           where (lvl.Name == levelName)
                           select lvl).First();
            return level;
        }

        /// <summary>
        /// find symbol of door.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="door"></param>
        /// <returns></returns>
        public FamilySymbol findSymbol(Document doc, ObjDoors door)
        {
            string fsFamilyName = "Door-Opening";
            string fsName = door.doorName;
            // LINQ to find the window's FamilySymbol by its type name.
            FamilySymbol familySymbol = (from fs in new FilteredElementCollector(doc).
                                             OfClass(typeof(FamilySymbol)).
                                             Cast<FamilySymbol>()
                                         where (fs.Family.Name == fsFamilyName && fs.Name == fsName)
                                         select fs).First();
            return familySymbol;
        }

        public static Element FindElementByName(Document doc, Type targetType, string targetName)
        {
            return new FilteredElementCollector(doc)
              .OfClass(targetType)
              .FirstOrDefault<Element>(
                e => e.Name.Equals(targetName));
        }
    }
}