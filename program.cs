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

            foreach (ObjDoors linkedDoor in List_DoorsLinkedModel)
            {
                // check to see if door exist 
                ObjDoors DoorFound = List_DoorsCurrentModel.Find(x => x.doorId == linkedDoor.doorId);
                
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
                }
            }

        }
        /// <summary>
        /// Create a door in the project 
        /// </summary>
        /// <param name="uidoc"></param>
        /// <param name="doc"></param>
        /// <param name="door"></param>
        public void CreateDoors(UIDocument uidoc, Document doc, ObjDoors door)
        {
            Level level = findLevel(doc, door);
            FamilySymbol familySymbol = findSymbol(doc, door);

            // Convert coordinates to double and create XYZ point.
            XYZ xyz = new XYZ(door.X, door.Y, door.Z);
           
            //Find the hosting Wall (nearst wall to the insertion point)
            //FilteredElementCollector collector = new FilteredElementCollector(doc);
            //collector.OfClass(typeof(Wall));
            //List<Wall> walls = collector.Cast<Wall>().Where(wl => wl.Id == level.Id).ToList();
            MyLibrary Library = new MyLibrary();
            List<Wall> walls = Library.GetWalls(doc);
            
            Wall wall = null;
            double distance = double.MaxValue;
            foreach (Element e in walls)
            {
                Wall w = e as Wall;
                LocationCurve lc = w.Location as LocationCurve;
                Curve curve = lc.Curve;
                XYZ z = curve.GetEndPoint(0);
                if (door.level.Elevation <= z.Z + 3 || door.level.Elevation >= z.Z - 3)
                {
                    int i = w.Id.IntegerValue;
                }
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
                if (!familySymbol.IsActive)
                {
                    // Ensure the family symbol is activated.
                    familySymbol.Activate();
                    doc.Regenerate();
                }
                // Create window
                // unliss you specified a host, Rebit will create the family instance as orphabt object.
                FamilyInstance fm = doc.Create.NewFamilyInstance(xyz, familySymbol, wall, level, StructuralType.NonStructural);
                Parameter parameter = fm.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                parameter.Set(door.doorId);
                t.Commit();
            }
        }

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
    }
}