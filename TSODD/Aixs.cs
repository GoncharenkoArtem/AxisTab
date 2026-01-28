using ACAD_test;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

namespace TSODD
{

    public class Axis
    {
        public Handle PolyHandle;
        public ObjectId PolyID;

        public Polyline AxisPoly;
        public string Name;
        //public Point3d StartPoint;

        public bool ReverseDirection = false;
        public double StartPK = 0.0;






        public bool GetAxisStartPoint()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            // Настройки промпта выбора начальной точки
            var peo = new PromptPointOptions("\n Выберете начальную точку оси: ");
            var per = ed.GetPoint(peo);
            if (per.Status != PromptStatus.OK || (AxisPoly == null))
            {
                ed.WriteMessage("\n Ошибка выбора начальной точки оси...");
                return false;
            }

            // поиск ближайшей вершины 
            double dist_1, dist_2;
            dist_1 = per.Value.DistanceTo(AxisPoly.StartPoint);
            dist_2 = per.Value.DistanceTo(AxisPoly.EndPoint);

            if (dist_1 < dist_2)
            {
                //StartPoint = AxisPoly.StartPoint;
                ReverseDirection = false;
            }
            else
            {
                //StartPoint = AxisPoly.EndPoint;
                ReverseDirection = true;
            }

            //ed.WriteMessage("\n Точка начала оси: X = " + StartPoint.X + ", Y = " + StartPoint.Y);

            // Настройки промпта выбора начальной точки
            var peoPk = new PromptDoubleOptions("\n Введите начальный пикет оси в метрах (1ПК = 100м): ");
            var perPk = ed.GetDouble(peoPk);
            if (perPk.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n Ошибка пикета оси...");
                return false;
            }

            StartPK = perPk.Value;

            return true;
        }


        public bool GetAxisName()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            // Настройки промпта
            var peo = new PromptStringOptions("\n Введите наименование оси: ")
            {
                AllowSpaces = true
            };

            var per = ed.GetString(peo);
            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n Ошибка ввода наименования оси...");
                return false;
            }

            Name = per.StringResult;
            ed.WriteMessage($"\n Новое имя оси: {per.StringResult}");
            return true;
        }


        public bool GetAxisPolyine()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            // Настройки промпта
            var peo = new PromptEntityOptions("\n Выберите полилинию: ");
            peo.SetRejectMessage("\n Это не полилиния. Выберите полилинию!");
            peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), exactMatch: false);

            // Запрос
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n Ошибка выбора полилинии...");
                return false;
            }

            // Работа с выбранной полилинией
            using (var tr = doc.TransactionManager.StartTransaction())
            {

                // приводим полилинию к простой линии
                RibbonInitializer.MLineTypeToPolyline(per.ObjectId, out var objId);

                var poly = tr.GetObject(objId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
                if (poly != null)
                {
                    PolyHandle = poly.Handle;
                    PolyID = poly.ObjectId;
                    AxisPoly = poly;
                }
                tr.Commit();
            }
            return true;
        }

    }
}
