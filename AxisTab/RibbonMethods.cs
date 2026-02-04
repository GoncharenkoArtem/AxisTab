using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media.Imaging;
using AxisTab;


/* Методы и обработчики для Ribbon */
namespace AxisTAb
{
    public partial class RibbonInitializer : IExtensionApplication
    {

        // обработчик события активации документа
        private void Dm_DocumentActivated(object sender, DocumentCollectionEventArgs e)
        {
            if (e?.Document == null) return;
            var db = e.Document.Database;
            if (db == null) return;
            var key = db.UnmanagedObject; // уникальный ключ DB

            if (_dbIntPtr.Add(key))       // если в структуре баз не было такого ключа, то добавит его
            {
                db.ObjectErased += Db_ObjectErased;     // событие на удаление объекта из DB]

                // события для отслеживания действий пользователя
                e.Document.Editor.LeavingQuiescentState += Editor_LeavingQuiescentState;
                e.Document.ViewChanged += Document_ViewChanged;
                e.Document.CommandEnded += MdiActiveDocument_CommandEnded;
                e.Document.CommandCancelled += MdiActiveDocument_CommandEnded;


            }
            
            ListOFAxisRebuild(); // перестраиваем combobox с осями

        }




        private void Document_ViewChanged(object sender, EventArgs e)
        {
            RestartTimer();
        }

        private void Editor_LeavingQuiescentState(object sender, EventArgs e)
        {
            RestartTimer();
        }

  
        // обработчик события закрытия документа
        private void Dm_DocumentToBeDestroyed(object sender, DocumentCollectionEventArgs e)
        {

            if (e?.Document == null) return;
            var db = e.Document.Database;
            if (db == null) return;
            var dockey = e.Document.UnmanagedObject;
            var dbkey = db.UnmanagedObject;

            if (_dbIntPtr.Remove(dbkey))
            {
                e.Document.CommandEnded -= MdiActiveDocument_CommandEnded;
                e.Document.CommandCancelled -= MdiActiveDocument_CommandEnded;
                e.Document.Editor.LeavingQuiescentState -= Editor_LeavingQuiescentState;
                e.Document.ViewChanged -= Document_ViewChanged;
                db.ObjectErased -= Db_ObjectErased;
            }

            DrawingHost.drawingDictionary.Remove(dockey);
        }




        // обработчик удаления элемента из базы autoCAD
        private void Db_ObjectErased(object sender, Autodesk.AutoCAD.DatabaseServices.ObjectErasedEventArgs e)
        {
           
            if (!e.Erased) return;
            if (e.DBObject.GetType().Name != "Polyline") return;

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;


            // проверяем есть ли xData у удаляемого объекта
            if (e.DBObject.XData != null)
            {
                //XdataElement tsoddXdataElement = new XdataElement();
                //try { tsoddXdataElement.Parse(e.DBObject.Id); }
                //catch (Autodesk.AutoCAD.Runtime.Exception ex) { return; }

                //if (tsoddXdataElement.Type == TsoddElement.Axis)    // если это ось
                //{
                //    // проверка является ли удаляемый объект осью
                //    var erasedObject = DrawingHost.Current.axis.FirstOrDefault(a => a.PolyID == e.DBObject.Id);

                //    if (erasedObject != null)
                //    {
                //        string msg = $"Полилиния привязана к оси {erasedObject.Name}. Вы точно хотите ее удалить?";

                //        // последнее предупреждение
                //        var result = MessageBox.Show(msg, "Сообщение", MessageBoxButton.YesNo, MessageBoxImage.Question);

                //        if (result == MessageBoxResult.No)
                //        {

                //            _dontDeleteMe.Add(erasedObject.PolyID);
                //        }
                //        else    // убираем из списка осей (удаляем невозвратно)
                //        {

                //            DrawingHost.Current.axis.Remove(erasedObject);

                //            // перестраиваем combobox с осями
                //            ListOFAxisRebuild();

                //        }
                //    }
                //}
                //else      // если это линия или мультилиния
                //{
                //    if (tsoddXdataElement.MasterPolylineID != ObjectId.Null) _deleteMe.Add(tsoddXdataElement.MasterPolylineID);
                //    if (tsoddXdataElement.SlavePolylineID != ObjectId.Null) _deleteMe.Add(tsoddXdataElement.SlavePolylineID);
                //    if (tsoddXdataElement.MtextID != ObjectId.Null) _deleteMe.Add(tsoddXdataElement.MtextID);
                //}
            }
        }


        // lifeguard для осей. после выполнения команды, проверяем есть ли в структуре _dontDeleteMe элементы, которые нужно спасти
        private void MdiActiveDocument_CommandEnded(object sender, CommandEventArgs e)
        {
            if (_dontDeleteMe.Count == 0 && _deleteMe.Count == 0) return;

            var doc = sender as Autodesk.AutoCAD.ApplicationServices.Document;
            var db = doc.Database;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // если есть элементы, которые нужно восстановить
                foreach (ObjectId id in _dontDeleteMe)
                {
                    var dbo = (DBObject)tr.GetObject(id, OpenMode.ForWrite, openErased: true);
                    if (dbo != null && dbo.IsErased)
                    {
                        // “unErase”
                        dbo.Erase(false);
                    }
                }

                // если есть элементы которые нужно удалить
                foreach (ObjectId id in _deleteMe)
                {
                    var dbo = (DBObject)tr.GetObject(id, OpenMode.ForWrite, openErased: true);
                    if (dbo != null && !dbo.IsErased)
                    {
                        // “Erase”
                        AutocadXData.UpdateXData(id, null);
                        dbo.Erase(true);
                    }
                }

                tr.Commit();
            }
            _dontDeleteMe.Clear();
            _deleteMe.Clear();

            RestartTimer();
        }


        // метод выбора и проверки полилинии оси
        public Axis SelectAxis()
        {
            Axis resultAxis = null;
            var doc = DrawingHost.Current.doc;
            var ed = doc.Editor;

            // Настройки промпта
            var peo = new PromptEntityOptions("\n Выберите ось: ");
            peo.SetRejectMessage("\n Это не полилиния. Выберите полилинию!");
            peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), exactMatch: false);

            // Запрос
            var per = ed.GetEntity(peo);
            if (per.Status == PromptStatus.OK)
            {
                resultAxis = DrawingHost.Current.axis.FirstOrDefault(a => a.PolyID == per.ObjectId);
                if (resultAxis == null)
                {
                    EditorMessage("\n Данная полилиния не является осью ");
                    return null;
                }
            }

            return resultAxis;
        }






        // метод перестраивания combobox осей
        public void ListOFAxisRebuild(Axis preselect = null)
        {
            if (axisCombo == null) return;
            axisCombo.Items.Clear();

            foreach (var a in DrawingHost.Current.axis)
                axisCombo.Items.Add(new RibbonButton { Text = a.Name, ShowText = true });

            // выбрать по умолчанию
            var toSelect = preselect != null
                ? axisCombo.Items.OfType<RibbonButton>().FirstOrDefault(b => b.Text == preselect.Name)
                : axisCombo.Items.OfType<RibbonButton>().FirstOrDefault();

            if (toSelect != null)
                axisCombo.Current = toSelect;

            if (preselect != null) DrawingHost.Current.currentAxis = preselect;
        }








        // обработчик события выбора текущей оси
        private void AxisCombo_CurrentChanged(object sender, RibbonPropertyChangedEventArgs e)
        {
            var curButton = axisCombo.Current as RibbonButton;
            if (curButton == null) return;
            DrawingHost.Current.currentAxis = DrawingHost.Current.axis.FirstOrDefault(a => a.Name == curButton.Text);
        }




        // метод получает объекты Axis с чертежа
        public static List<Axis> GetListOfAxis()
        {
            List<Axis> list = new List<Axis>();

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var filter = new SelectionFilter(new TypedValue[]
                { new TypedValue((int)DxfCode.Start,"LWPOLYLINE,POLYLINE"),
                  new TypedValue((int)DxfCode.ExtendedDataRegAppName,AutocadXData.AppName),
                  new TypedValue((int)DxfCode.ExtendedDataAsciiString,"axis")
                }
                );

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {

                // поиск с учетом фильтра
                var selection = ed.SelectAll(filter);
                if (selection.Status != PromptStatus.OK)    // неудачный поиск
                {
                    ed.WriteMessage("\n В чертеже не найдено осей \n");
                    return list;
                }
                else     // удачный поиск
                {
                    foreach (var id in selection.Value.GetObjectIds())  // проходимся по всем объектам и читам XData
                    {
                        // создаем экземпляр Axis
                        Axis axis = AutocadXData.AxisXdataParse(id);
                        
                        // добавляем ось в список осей
                        list.Add(axis);
                    }
                }
            }
            return list;
        }





        public List<ObjectId> GetAutoCadSelectionObjectsId(SelectionFilter filter, string textMessage,  bool singleSelection = false)
        {
            var doc = DrawingHost.Current.doc;
            var ed = doc.Editor;

            var pso = new PromptSelectionOptions
            {
                MessageForAdding = textMessage,
                SingleOnly = singleSelection,
                AllowDuplicates = false
            };

            var psr = ed.GetSelection(pso, filter);

            if (psr.Status != PromptStatus.OK) return null; // неудачный промпт, выходим

            List<ObjectId> resultList = new List<ObjectId>();
            foreach (var id in psr.Value.GetObjectIds()) resultList.Add(id);

            return resultList;
        }



        public List<ObjectId> GetAutoCadSelectionObjectsId(List<string> filterList, string textMessage, bool singleSelection = false)
        {
            // массив TypedValue[] для фильтра
            TypedValue[] typedValue = new TypedValue[filterList.Count + 2];
            SelectionFilter filter = null;
            if (filterList.Count > 0)
            {
                filterList = filterList.Select(a => a.ToUpper()).ToList(); // все в верхний регистр
                typedValue[0] = new TypedValue((int)DxfCode.Operator, "<OR");
                typedValue[typedValue.Length - 1] = new TypedValue((int)DxfCode.Operator, "OR>");
                for (int i = 1; i < typedValue.Length - 1; i++) typedValue[i] = new TypedValue((int)DxfCode.Start, filterList[i - 1]);
                filter = new SelectionFilter(typedValue);
            }

            return GetAutoCadSelectionObjectsId(filter, textMessage, singleSelection);
        }


        // метод подгружающий картинки для кнопок
        public BitmapImage LoadImage(string uri)
        {
            try
            {
                BitmapImage image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri(uri, UriKind.RelativeOrAbsolute);
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.EndInit();
                image.Freeze(); // Важно для многопоточности
                return image;
            }
            catch
            {
                // Возвращаем пустое изображение при ошибке
                return new BitmapImage();
            }
        }


        // вывод сообщения в editor
        private static void EditorMessage(string txt)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            ed.WriteMessage($" \n {txt}");
        }

        /// <summary>
        ///  Получает ПК оси указанной точки
        /// </summary>
        /// 
        public void GetPkOnAxis()
        {
            var doc = DrawingHost.Current.doc;
            var db = doc.Database;
            var ed = doc.Editor;

            // выбор оси
            Axis axis = RibbonInitializer.Instance?.SelectAxis();
            if (axis == null) return; // выходим, ошибка выбора оси
            var polylineAxis = axis.AxisPoly;
            if (polylineAxis == null) return; // выходим, ошибка выбора оси
            double koef = db.Insunits == UnitsValue.Meters ? 1 : 0.001;   // коэффициент перевода в метры
            string PK = string.Empty;
            Point3d cursorPos = Point3d.Origin;

            // Подписываемся на событие
            ed.PointMonitor += OnPointMonitor;

            try
            {
                var peo = new PromptPointOptions("\n Точка (x,y):");
                var per = ed.GetPoint(peo);
                if (per.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\n Ошибка подсчета ПК \n");
                    return;
                }
            }
            finally
            {
                ed.PointMonitor+= OnPointMonitor;
                if (string.IsNullOrEmpty(PK))
                {
                    ed.WriteMessage("\n Ошибка подсчета ПК \n");
                }
                else
                {
                    ed.WriteMessage($"{cursorPos.X},{cursorPos.Y}; {PK}; Ось: {axis.Name} \n");
                }
            }

            // обработчик курсора
            void OnPointMonitor(object sender, PointMonitorEventArgs e)
            {
                cursorPos = e.Context.ComputedPoint;    // Получаем текущую позицию курсора
                Point3d textPoint = new Point3d(cursorPos.X + 150, cursorPos.Y + 150, 0);
                Point3d pointOnAxis = polylineAxis.GetClosestPointTo(cursorPos, false);     // точка на полилинии оси

                // Используем DrawContext для рисования
                var dc = e.Context.DrawContext;
                if (dc != null)
                {
                    Point3dCollection pointsCollection = new Point3dCollection { cursorPos, pointOnAxis };
                    Autodesk.AutoCAD.GraphicsInterface.Polyline tempPoly = new Autodesk.AutoCAD.GraphicsInterface.Polyline(pointsCollection, Vector3d.ZAxis, IntPtr.Zero);
                    dc.Geometry.Polyline(tempPoly);
                    tempPoly.Dispose();

                    // считаем ПК
                    var distance = polylineAxis.GetDistAtPoint(pointOnAxis) * koef;   // расстояние от начала 
                    if (axis.ReverseDirection) distance = polylineAxis.Length * koef - distance;    // если реверсивное направление оси
                    distance = Math.Round(distance, 3) + axis.StartPK * 100;

                    int pt_1 = (int)Math.Truncate(distance / 100);
                    double pt_2 = Math.Round((distance - pt_1 * 100), 2);
                    PK = $"ПК {pt_1} + {pt_2}";

                    e.AppendToolTipText(""); // Очистка
                    e.AppendToolTipText(PK); // пикет
                }
            }
        }


        /// <summary>
        /// Отстраиывает ПК на оси, по указанному double
        /// </summary>
        public void SetPkOnAxis()
        {
            var doc = DrawingHost.Current.doc;
            var db = doc.Database;
            var ed = doc.Editor;

            // выбор оси
            Axis axis = RibbonInitializer.Instance?.SelectAxis();
            if (axis == null) return; // выходим, ошибка выбора оси
            var polylineAxis = axis.AxisPoly;
            if (polylineAxis == null) return; // выходим, ошибка выбора оси
            double koef = db.Insunits == UnitsValue.Meters ? 1 : 0.001;   // коэффициент перевода в метры

            var peoPk = new PromptDoubleOptions("\n Введите пикет в метрах (1ПК = 100м): ");
            var perPk = ed.GetDouble(peoPk);
            if (perPk.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n Ошибка пикета оси...");
                return;
            }

            double distVal = perPk.Value;

            // проверка значения на экстремумы
            if (axis.StartPK > distVal || distVal > (axis.AxisPoly.Length / 100) * koef + axis.StartPK)
            {
                ed.WriteMessage("\n Значение выходит за пределы значений ПК оси \n");
                return;
            }

            // Находим точку на оси и отстраиваем линию
            var distance = distVal * 100 / koef - axis.StartPK * 100 / koef;
            Point3d pointOnAxis = axis.ReverseDirection ? axis.AxisPoly.GetPointAtDist(axis.AxisPoly.Length - distance) : axis.AxisPoly.GetPointAtDist(distance);

            double parametr = axis.AxisPoly.GetParameterAtPoint(pointOnAxis);
            Vector3d vector = axis.AxisPoly.GetFirstDerivative(parametr).GetNormal().GetPerpendicularVector();
            double angle = Math.Atan2(vector.Y, vector.X);


            var userOptions = JsonReader.LoadFromJson<Options>(FilesLocation.JsonOptionsPath);
            var distanceOffsetFromAxis = userOptions.PKLineLength / koef;
            Point3d p_1 = pointOnAxis + vector * distanceOffsetFromAxis;
            Point3d p_2 = pointOnAxis - vector * distanceOffsetFromAxis;

            int pt_1 = (int)Math.Truncate(distVal);
            double pt_2 = Math.Round((distVal - pt_1) * 100, 2);
            string PK = $"ПК {pt_1} + {pt_2}";

            Line line = new Line(p_1, p_2);
            line.Layer = "0";

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ocm = db.ObjectContextManager;
                var occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");

                // Текущий масштаб CANNOSCALE
                var curSC = occ.CurrentContext as AnnotationScale;

                // Текст ПК
                MText mtext = new MText();
                mtext.Contents = PK;
                mtext.Rotation = angle;
                mtext.TextHeight = userOptions.PKTextHeight / curSC.Scale;
                mtext.Location = pointOnAxis + ((vector - vector.GetPerpendicularVector()) / curSC.Scale);

                // Получаем таблицу текстовых стилей
                TextStyleTable textStyleTable = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);

                // Перебираем все стили
                foreach (ObjectId styleId in textStyleTable)
                {
                    TextStyleTableRecord styleRecord = (TextStyleTableRecord)tr.GetObject(styleId, OpenMode.ForRead);
                    if (!string.IsNullOrEmpty(styleRecord.Name) && styleRecord.Name == userOptions.PKTextStyle) mtext.TextStyleId = styleId;
                }

                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                btr.AppendEntity(mtext);
                tr.AddNewlyCreatedDBObject(mtext, true);
                btr.AppendEntity(line);
                tr.AddNewlyCreatedDBObject(line, true);

                tr.Commit();
            }
        }











    }
}
