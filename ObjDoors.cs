using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SOM.RevitTools.PlaceDoors
{
    class ObjDoors
    {
        public Element doorElement { get; set; }
        public string doorName{ get; set; }
        public string doorId { get; set; }
        public double doorHeight { get; set; }
        public double doorWidth { get; set; }
        public double doorSillHeight { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public Level level { get; set; }
        public int HostObj { get; set; }
    }
}
