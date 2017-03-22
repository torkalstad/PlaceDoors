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
//using _00_RevitLibrary;
using DBLibrary;
using NLog;

namespace SOM.RevitTools.PlaceDoors
{
    class program
    {
        //*****************************logger()*****************************
        private static Logger logger = LogManager.GetCurrentClassLogger();

        private static List<Level> _ListLevels = null;
        //*****************************DoorProgram()*****************************
        public List<ObjDoors> DoorProgram(Document doc, UIDocument uidoc,
            List<ObjDoors> List_DoorsLinkedModel, List<ObjDoors> List_DoorsCurrentModel)
        {            
            List<ObjDoors> ListObjDoors = new List<ObjDoors>();
            // logger object. 
            Logger logger = LogManager.GetLogger("program");
            ExportExcel exportExcel = new ExportExcel();

            // Get all levels and add to field. 
            LibraryGetItems libGetItems = new LibraryGetItems();
            _ListLevels = libGetItems.GetLevels(doc);

            foreach (ObjDoors linkedDoor in List_DoorsLinkedModel)
            {
                
                // check to see if door exist 
                ObjDoors DoorFound = List_DoorsCurrentModel.Find(x => x.doorId == linkedDoor.doorId);
                // if it doesn't exist it will create a new door. 
                if (DoorFound == null)
                {
                    try
                    {
                        ObjDoors door = CreateDoors(uidoc, doc, linkedDoor);
                        ListObjDoors.Add(door);
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

                    // Get all the levels from current project. 
                    Level level = findLevel(doc, linkedDoor);

                    // Level offset from architecture. 
                    LibraryGetItems Library = new LibraryGetItems();
                    LibraryConvertUnits library_Units = new LibraryConvertUnits();
                    double offset = 0;
                    Level CurrentLevel = _ListLevels.Find(x => x.Name == level.Name);
                    try
                    {

                        Parameter parameter_LevelOffsetFF = CurrentLevel.LookupParameter("LEVEL OFFSET FF");
                        offset = library_Units.InchesToFeet(Convert.ToDouble(Library.GetParameterValue(parameter_LevelOffsetFF)));
                    }
                    catch { }

                    // Check to see if door size match. 
                    double height = Math.Round(linkedDoor.doorHeight + offset, 3);
                    double width = Math.Round(linkedDoor.doorWidth, 3);
                    String doorType = doorTypeByUnits(doc, height, width);

                    if (doorType != "0 x 0")
                    {
                        if (DoorFound.doorName != doorType)
                        {
                            //FamilySymbol familySymbol = findSymbol(doc, DoorFound, doorType);
                            FamilySymbol familySymbol = FindElementByName(doc, typeof(FamilySymbol), doorType) as FamilySymbol;
                            if (familySymbol == null)
                            {
                                FamilySymbol oldType = findSymbol(doc, DoorFound);

                                FamilySymbol ChangeFamilySymbol = CreateNewType(doc, oldType, linkedDoor, offset);
                                //FamilySymbol ChangeFamilySymbol = BKK_CreateNewType(doc, oldType, linkedDoor);
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
            //REMOVE UNSUED DOORS. 
            RemoveUnused(doc, List_DoorsLinkedModel, List_DoorsCurrentModel);

            return ListObjDoors;
        }

        #region CREATE ELEMENTS

        //*****************************CreateDoors()*****************************
        public ObjDoors CreateDoors(UIDocument uidoc, Document doc, ObjDoors linkedDoor)
        {
            LibraryGetItems Library = new LibraryGetItems();
            LibraryConvertUnits library_Units = new LibraryConvertUnits();
            // Get all the levels from current project. 
            Level level = findLevel(doc, linkedDoor);

            // Level offset from architecture. 
            double offset = 0;
            Level CurrentLevel = _ListLevels.Find(x => x.Name == level.Name);
            try
            {
                
                Parameter parameter_LevelOffsetFF = CurrentLevel.LookupParameter("LEVEL OFFSET FF");
                offset = library_Units.InchesToFeet(Convert.ToDouble(Library.GetParameterValue(parameter_LevelOffsetFF)));                
            }
            catch { }
            

            // CREATE : door type height x width 
            double height = Math.Round(linkedDoor.doorHeight + offset, 3);
            double width = Math.Round(linkedDoor.doorWidth, 3);
            String doorType = doorTypeByUnits(doc, height, width);

            if (doorType != "0 x 0")
            {
                FamilySymbol currentDoorType = null;
                try
                {
                    currentDoorType = FindElementByName(doc, typeof(FamilySymbol), doorType) as FamilySymbol;
                }
                catch { }

                if (currentDoorType == null)
                {
                    FamilySymbol familySymbol_OldType = FindElementByName(doc, typeof(FamilySymbol), "10 x 10") as FamilySymbol;
                    currentDoorType = CreateNewType(doc, familySymbol_OldType, linkedDoor, offset);
                }

                // Convert coordinates to double and create XYZ point. Use offset level from current model.
                XYZ xyz = new XYZ(linkedDoor.X, linkedDoor.Y, CurrentLevel.Elevation + linkedDoor.doorSillHeight);

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

                        var SOMIDParam = e.LookupParameter("SOM ID");
                        string wallSOMId = Library.GetParameterValue(SOMIDParam);
                        if (linkedDoor.HostObj.ToString() == wallSOMId)
                        {
                            wall = w;
                        }
                            
                    }
                }
                try
                {
                    if (wall != null)
                    {
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
                            // unless you specified a host, Revit will create the family instance as orphabt object.
                            FamilyInstance fm = doc.Create.NewFamilyInstance(xyz, currentDoorType, wall, level, StructuralType.NonStructural);
                            // Set new local door id to match linked model element id. 
                            Parameter SOMIDParam = fm.LookupParameter("SOM ID");
                            SOMIDParam.Set(linkedDoor.doorId);
                            linkedDoor.wall = wall;
                            t.Commit();
                            return linkedDoor;
                        }
                    }
                }
                catch { }
            }
            return null;
        }

        //*****************************CreateNewType()*****************************
        public FamilySymbol CreateNewType(Document doc, FamilySymbol oldType, ObjDoors linkedDoor, double offset)
        {
            double height = Math.Round(linkedDoor.doorHeight + offset, 3);
            double width = Math.Round(linkedDoor.doorWidth, 3);
            String doorType = doorTypeByUnits(doc, height, width);

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

        #endregion


        #region MOVE, CHANGE TYPE AND REMOVE UNUSED 

        //*****************************MoveDoors()*****************************
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

        //*****************************RemoveUnused()*****************************
        public void RemoveUnused(Document doc, List<ObjDoors> List_DoorsLinkedModel, List<ObjDoors> List_DoorsCurrentModel)
        {
            foreach (ObjDoors CurrentDoor in List_DoorsCurrentModel)
            {
                // check to see if door exist 
                ObjDoors DoorFound = List_DoorsLinkedModel.Find(x => x.doorId == CurrentDoor.doorId);
                // if it doesn't exist it will remove door from project
                if (DoorFound == null)
                {
                    try
                    {
                        Transaction t = new Transaction(doc);
                        t.Start("Remove Element");
                        Element e = CurrentDoor.doorElement;
                        doc.Delete(e.Id);
                        t.Commit();
                    }
                    catch (Exception e)
                    {
                        string ErrMessage = e.Message;
                    }
                }
            }
        }

        //*****************************changeType()*****************************
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
        #endregion


        #region LEVELS, SYMBOLS AND ELEMENT NAMES

        //*****************************doorTypeByUnits()*****************************
        public string doorTypeByUnits(Document doc, double height, double width)
        {
            LibraryConvertUnits lib = new LibraryConvertUnits();
            string doorType = "";
            double h = 12 * height;
            double w = 12 * width;
            doorType = h.ToString() + " x " + w.ToString();

            return doorType;
        }

        //*****************************findLevel()*****************************
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

        //*****************************findSymbol()*****************************
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

        //*****************************FindElementByName()*****************************
        public static Element FindElementByName(Document doc, Type targetType, string targetName)
        {
            return new FilteredElementCollector(doc)
              .OfClass(targetType)
              .FirstOrDefault<Element>(
                e => e.Name.Equals(targetName));
        }

        //*****************************GetDesignOption()*****************************
        public bool GetDesignOption(Document doc, Element e)
        {
            bool result = true;
            if (e.DesignOption != null)
            {
                if (e.DesignOption.Name == "T7 Basement - Opt 1 - under Bar (primary)")
                    result = true;
                else
                    result = false;
            }
            return result;
        }

        #endregion
    }
}