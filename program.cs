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
        //*****************************_List_CreatedDoors()*****************************
        public List<ObjDoors> _List_CreatedDoors {get; set;}

        //*****************************logger()*****************************
        private static Logger logger = LogManager.GetCurrentClassLogger();

        //*****************************DoorProgram()*****************************
        public void DoorProgram(Document doc, UIDocument uidoc,
            List<ObjDoors> List_DoorsLinkedModel, List<ObjDoors> List_DoorsCurrentModel)
        {
            MyLibrary lib = new MyLibrary();

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
                        //LACMA_CreateDoors(uidoc, doc, linkedDoor);
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
                    //double height = Math.Round(lib.FootToMm(linkedDoor.doorHeight), 2);
                    //double width = Math.Round(lib.FootToMm(linkedDoor.doorWidth), 2);
                    double height = Math.Round(linkedDoor.doorHeight, 2);
                    double width = Math.Round(linkedDoor.doorWidth, 2);
                    //String doorType = height.ToString() + "MM" + " x " + width.ToString() + "MM";
                    String doorType = doorTypeByUnits(doc, height, width);

                    //if (doorType != "0MM x 0MM")
                    if (doorType != "0 x 0")
                    {
                        if (DoorFound.doorName != doorType)
                        {
                            //FamilySymbol familySymbol = findSymbol(doc, DoorFound, doorType);
                            FamilySymbol familySymbol = FindElementByName(doc, typeof(FamilySymbol), doorType) as FamilySymbol;
                            if (familySymbol == null)
                            {
                                FamilySymbol oldType = findSymbol(doc, DoorFound);

                                FamilySymbol ChangeFamilySymbol = CreateNewType(doc, oldType, linkedDoor);
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
        }

        #region CREATE ELEMENTS FOR LACMA 

        //*****************************CreateDoors()*****************************
        public void CreateDoors(UIDocument uidoc, Document doc, ObjDoors linkedDoor)
        {
            MyLibrary Library = new MyLibrary();
            Level level = findLevel(doc, linkedDoor);
            // make door type height x width 
            double height = Math.Round(linkedDoor.doorHeight, 2);
            double width = Math.Round(linkedDoor.doorWidth, 2);
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
                        
                        var SOMIDParam = e.LookupParameter("SOM ID");
                        string wallSOMId = Library.GetParameterValue(SOMIDParam);
                        if (linkedDoor.HostObj.ToString() == wallSOMId)
                        {
                            wall = w;
                        }
                    }
                }
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
        }

        

        //*****************************CreateNewType()*****************************
        public FamilySymbol CreateNewType(Document doc, FamilySymbol oldType, ObjDoors linkedDoor)
        {
            double height = Math.Round(linkedDoor.doorHeight, 2);
            double width = Math.Round(linkedDoor.doorWidth, 2);
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

        #region find Revit levels, symbols and element names. 

        //*****************************doorTypeByUnits()*****************************
        public string doorTypeByUnits(Document doc, double height, double width)
        {
            MyLibrary lib = new MyLibrary();
            string doorType = "";

            if (doc.ProjectInformation.Number == "215152")
            {
                doorType = lib.FootToMm(height).ToString() + " x " + lib.FootToMm(width).ToString();
            }
            else
            {
                double h = 12 * height;
                double w = 12 * width;
                doorType = h.ToString() + " x " + w.ToString();
            }
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

        #endregion
    }
}