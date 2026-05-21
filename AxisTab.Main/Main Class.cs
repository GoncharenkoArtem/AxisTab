using System.Collections.Generic;

namespace AxisTab
{
    public class DrawingAxis
    {
        public Autodesk.AutoCAD.ApplicationServices.Document doc { get; set; }  // текущий документ
        public List<Axis> axis { get; set; } = new List<Axis>();    // список всех осей
        public Axis currentAxis { get; set; }                       // текущая ось
    }
}
