using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Geometry;
using System.Windows;
using Autodesk.AutoCAD.Runtime;
using ACAD_test;
using Autodesk.Windows;
using TSODD;

namespace TSODD
{
    public class DrawingTSODD
    {
        public List<Axis> axis { get; set; } = new List<Axis>();    // список всех осей
        public Axis currentAxis { get; set; }                       // текущая ось
        public string currentSignGroup { get; set; }                // текущая группа стоек
        public string currentStandBlock { get; set; }               // текущий блок стойки
        public string currentSignBlock { get; set; }                // текущий блок стойки
        public string currentMarksLineType { get; set; }            // текущий блок стойки
    }


    public interface IRefBlock 
    {
        string Tag { get; }                         // тег блока 
    }

    public class Stand : IRefBlock
    {
        public string Tag { get; } = "STAND";
        public string Handle { get; set; }          // handle стойки
        public string AxisName { get; set; }        // имя оси к которой привязана стойка
        public double Distance { get; set; }        // расстояние от начала оси до точки ПК (для сортировки)
        public string PK { get; set; }              // ПК в формате string, относительно привязанной оси
        public string Side { get; set; }            // слева или справа от оси
    }

    public class Sign : IRefBlock
    {
        public string Tag { get; } = "SIGN";
        public string StandHandle { get; set; }     // handle стойки, к которому привязан знак
        public string Number { get; set; }          // номер знака
        public string Name { get; set; }            // наименование знака
        public string TypeSize { get; set; }        // типоразмер знака
        public string Doubled { get; set; }           // сдвоенный (1 или 2 на одной стойке)
        public string Existence { get; set; }       // наличие ( Необходимо установить / Необходимо демонтировать)

    }

}
