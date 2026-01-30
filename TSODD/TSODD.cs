using Autodesk.AutoCAD.DatabaseServices;
using System.Collections.Generic;

namespace TSODD
{

    public class DrawingTSODD
    {
        public Autodesk.AutoCAD.ApplicationServices.Document doc { get; set; } = null;// текущий документ
        public List<Axis> axis { get; set; } = new List<Axis>();    // список всех осей
        public Axis currentAxis { get; set; }                       // текущая ось
        public string currentSignGroup { get; set; }                // текущая группа знаков
        public string currentMarkGroup { get; set; }                // текущая группа разметки
        public string currentStandBlock { get; set; }               // текущий блок стойки
        public string currentSignBlock { get; set; }                // текущий блок знака
        public string currentMarkBlock { get; set; }                // текущий блок разметки
        public string currentMarksLineType { get; set; }            // текущий блок стойки

    }


    public interface IRefBlock
    {
        ObjectId ID { get; set; }                   // ID блока
        string Tag { get; }                         // тег блока 
    }

    public class Stand : IRefBlock
    {
        public ObjectId ID { get; set; }            // ID
        public string Tag { get; } = "STAND";
        public string Handle { get; set; }          // handle стойки
        public string AxisName { get; set; }        // имя оси к которой привязана стойка
        public double Distance { get; set; }        // расстояние от начала оси до точки ПК (для сортировки)
        public string PK { get; set; }              // ПК в формате string, относительно привязанной оси
        public string Side { get; set; }            // слева или справа от оси
    }

    public class Sign : IRefBlock
    {
        public ObjectId ID { get; set; }            // ID
        public string Tag { get; } = "SIGN";
        public string StandHandle { get; set; }     // handle стойки, к которому привязан знак
        public string Number { get; set; }          // номер знака
        public string Name { get; set; }            // наименование знака
        public string TypeSize { get; set; }        // типоразмер знака
        public string Doubled { get; set; }         // сдвоенный (1 или 2 на одной стойке)
        public string Existence { get; set; }       // наличие ( Необходимо установить / Необходимо демонтировать)

    }

    public class Mark : IRefBlock
    {
        public ObjectId ID { get; set; }            // ID
        public string Tag { get; } = "MARK";
        public string AxisName { get; set; }        // имя оси к которой привязана разметка
        public string Number { get; set; }          // номер разметки
        public string Quantity { get; set; }        // количество
        public double Square { get; set; }          // приведенная площадь
        public string PK_start { get; set; }        // ПК в формате string, относительно привязанной оси (начало)
        public string PK_end { get; set; }          // ПК в формате string, относительно привязанной оси (конец)
        public double Distance { get; set; }        // расстояние от начала оси до точки ПК (для сортировки)
        public string Side { get; set; }            // слева или справа от оси
        public string Material { get; set; }        // материал разметки
        public string Existence { get; set; }       // наличие ( Необходимо установить / Необходимо демонтировать)

    }




}
