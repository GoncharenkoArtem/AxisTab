using AxisTAb;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Windows;

namespace AxisTAb
{

    public class Axis
    {
        public Handle PolyHandle;
        public ObjectId PolyID;

        public Polyline AxisPoly;
        public string Name;

        public bool ReverseDirection = false;
        public double StartPK = 0.0;


        // Метод создания новой оси 
        public void NewAxis()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            // заролняем экземпляр
            if (!this.GetAxisPolyine()) return;

            // проверка уникальности полилинии
            Axis dublicate = DrawingHost.Current.axis.FirstOrDefault(h => h.PolyHandle == this.PolyHandle);
            if (dublicate != null)
            {
                MessageBox.Show($" Ошибка создания новой оси. Данная полилиния уже была выбрана для оси \" {dublicate.Name} \" ");
                return;
            }

            if (!this.GetAxisName()) return;

            // проверка уникальности наименования оси
            dublicate = DrawingHost.Current.axis.FirstOrDefault(h => h.Name == this.Name);
            if (dublicate != null)
            {
                MessageBox.Show($" Ошибка создания новой оси. Ось с наименованием \" {dublicate.Name} \" уже существует.");
                return;
            }

            if (!this.GetAxisStartPoint()) return;

            // сообщение о удачном создании оси
            ed.WriteMessage($"\n Создана новая ось " +
                $"\n Имя оси: {this.Name} " +
                $"\n Начальный пикет: {Math.Round(this.StartPK, 3)} \n");

            // Записываем Xdata в полилинию оси
            List<(int, string)> xDataList = new List<(int, string)>();
            xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, "axis"));                                           // идентификатор линии (ось)
            xDataList.Add(((int)DxfCode.ExtendedDataHandle, this.PolyHandle.ToString()));                            // handle линии
            xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, $"{this.Name}"));                                   // имя оси

            xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, $"{this.ReverseDirection.ToString()}"));            // обратное или прямое направление
            xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, $"{this.StartPK.ToString()}"));                     // начальный ПК

            AutocadXData.UpdateXData(this.PolyID, xDataList);

            // добавляем ось в список осей
            DrawingHost.Current.axis.Add(this);

            // перестраиваем combobox наименование осей на ribbon
            RibbonInitializer.Instance?.ListOFAxisRebuild(this);
        }


        public bool GetAxisStartPoint()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            // xData оси
            var listXData = AutocadXData.ReadXData(this.PolyID);

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
                ReverseDirection = false;
                if (listXData.Count > 0) listXData[3] = ((int)DxfCode.ExtendedDataAsciiString, "false");
            }
            else
            {
                ReverseDirection = true;
                if (listXData.Count > 0) listXData[3] = ((int)DxfCode.ExtendedDataAsciiString, "true");
            }

            // Настройки промпта выбора начальной точки
            var peoPk = new PromptDoubleOptions("\n Введите начальный пикет оси в метрах (1ПК = 100м): ");
            var perPk = ed.GetDouble(peoPk);
            if (perPk.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n Ошибка пикета оси...");
                return false;
            }

            StartPK = perPk.Value;

            if (listXData.Count > 0)  // если это перезапись данных
            {
                listXData[4] = ((int)DxfCode.ExtendedDataAsciiString, $"{this.StartPK.ToString()}");
                // перезаписываем  xData оси
                AutocadXData.UpdateXData(this.PolyID, listXData);
            }

            return true;
        }

        public bool GetAxisName()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            // xData оси
            var listXData = AutocadXData.ReadXData(this.PolyID);


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
            
            if (listXData.Count>0)  // если это перезапись данных
            {
                listXData[2] = ((int)DxfCode.ExtendedDataAsciiString, $"{this.Name}");
                // перезаписываем  xData оси
                AutocadXData.UpdateXData(this.PolyID, listXData);
            }

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
                //RibbonInitializer.MLineTypeToPolyline(per.ObjectId, out var objId);

                var poly = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
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
