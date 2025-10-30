using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System.Windows.Input;
using System.Linq;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using TSODD;
using Autodesk.AutoCAD.ApplicationServices;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using System.Reflection;
using System;
using System.Windows.Documents;
using System.Xml.Linq;
using Autodesk.AutoCAD.EditorInput;
using System.Security.Cryptography;
using System.Windows.Controls;
using System.Drawing;
using System.Windows.Media.Media3D;
using System.Security.Policy;
using Autodesk.AutoCAD.Geometry;




/* Методы и обработчики для Ribbon */


namespace ACAD_test
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
                e.Document.CommandEnded += MdiActiveDocument_CommandEnded;
                e.Document.CommandCancelled += MdiActiveDocument_CommandEnded;
                db.ObjectErased += Db_ObjectErased;     // событие на удаление объекта из DB]
                //db.BeginSave += Db_BeginSave; ;
            }

            // перестраиваем combobox с осями
            ListOFAxisRebuild();

            TsoddBlock.blockInsertFlag = true; // разрешаем вставку блоков
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
                db.ObjectErased -= Db_ObjectErased;
            }

            TsoddHost.tsoddDictionary.Remove(dockey);

        }





        // обработчик сохранения файлов
        //private void Db_BeginSave(object sender, DatabaseIOEventArgs e)
        //{
        //    var file = e.FileName;
        //    var extension = file.Split('.')[1];
        //    bool autoSave = extension.Contains("$");

        //    // если это не автосейв, то обновим данные 
        //    if (!autoSave)
        //    {
        //        SaveTsoddChanges();
        //    }
        //}


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

                TsoddXdataElement tsoddXdataElement = new TsoddXdataElement();
                tsoddXdataElement.Parse(e.DBObject.Id);

                if (tsoddXdataElement.Type == TsoddElement.Axis)    // если это ось
                {

                    // проверка является ли удаляемый объект осью
                    var erasedObject = TsoddHost.Current.axis.FirstOrDefault(a => a.PolyID == e.DBObject.Id);

                    if (erasedObject != null)
                    {
                        string msg = $"Полилиния привязана к оси {erasedObject.Name}. Вы точно хотите ее удалить?";

                        // последнее предупреждение
                        var result = MessageBox.Show(msg, "Сообщение", MessageBoxButton.YesNo, MessageBoxImage.Question);

                        if (result == MessageBoxResult.No)
                        {

                            _dontDeleteMe.Add(erasedObject.PolyID);
                        }
                        else    // убираем из списка осей (удаляем невозвратно)
                        {

                            TsoddHost.Current.axis.Remove(erasedObject);

                            // перестраиваем combobox с осями
                            ListOFAxisRebuild();

                        }
                    }
                }
                else      // если это линия ьили мультилиния
                {
                    if (tsoddXdataElement.MasterPolylineID != ObjectId.Null) _deleteMe.Add(tsoddXdataElement.MasterPolylineID);
                    if (tsoddXdataElement.SlavePolylineID != ObjectId.Null) _deleteMe.Add(tsoddXdataElement.SlavePolylineID);
                    if (tsoddXdataElement.MtextID != ObjectId.Null) _deleteMe.Add(tsoddXdataElement.MtextID);
                }
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
        }



        // метод перестраивания combobox осей
        public void ListOFAxisRebuild(Axis preselect = null)
        {
            if (axisCombo == null) return;
            axisCombo.Items.Clear();

            foreach (var a in TsoddHost.Current.axis)
                axisCombo.Items.Add(new RibbonButton { Text = a.Name, ShowText = true });

            // выбрать по умолчанию
            var toSelect = preselect != null
                ? axisCombo.Items.OfType<RibbonButton>().FirstOrDefault(b => b.Text == preselect.Name)
                : axisCombo.Items.OfType<RibbonButton>().FirstOrDefault();

            if (toSelect != null)
                axisCombo.Current = toSelect;

            if(preselect != null) TsoddHost.Current.currentAxis = preselect;
        }



        // Метод создания новой оси 
        public void NewAxis()
        {
            // экземпляр новой оси
            Axis newAxis = new Axis();

            // заролняем экземпляр
            if (!newAxis.GetAxisPolyine()) return;

            // проверка уникальности полилинии
            Axis dublicate = TsoddHost.Current.axis.FirstOrDefault(h => h.PolyHandle == newAxis.PolyHandle);
            if (dublicate != null)
            {
                MessageBox.Show($" Ошибка создания новой оси. Данная полилиния уже была выбрана для оси \" {dublicate.Name} \" ");
                return;
            }

            if (!newAxis.GetAxisName()) return;

            // проверка уникальности наименования оси
            dublicate = TsoddHost.Current.axis.FirstOrDefault(h => h.Name == newAxis.Name);
            if (dublicate != null)
            {
                MessageBox.Show($" Ошибка создания новой оси. Ось с наименованием \" {dublicate.Name} \" уже существует.");
                return;
            }

            if (!newAxis.GetAxisStartPoint()) return;

            // сообщение о удачном создании оси
            EditorMessage($"\n Создана новая ось " +
                $"\n Имя оси: {newAxis.Name} " +
                $"\n Начальный пикет: {Math.Round(newAxis.StartPK, 3)} \n");

            // Записываем Xdata в полилинию оси
            List<(int, string)> xDataList = new List<(int, string)>();
            xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, "axis"));                                              // идентификатор линии (ось)
            xDataList.Add(((int)DxfCode.ExtendedDataHandle, newAxis.PolyHandle.ToString()));                            // handle линии
            xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, $"{newAxis.Name}"));                                   // имя оси
            xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, $"{newAxis.ReverseDirection.ToString()}"));            // обратное или прямое направление
            xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, $"{newAxis.StartPK.ToString()}"));                     // начальный ПК

            AutocadXData.UpdateXData(newAxis.PolyID, xDataList);

            // добавляем ось в список осей
            TsoddHost.Current.axis.Add(newAxis);

            // перестраиваем combobox наименование осей на ribbon
            RibbonInitializer.Instance?.ListOFAxisRebuild(newAxis);
        }





        // обработчик события выбора текущей оси
        private void AxisCombo_CurrentChanged(object sender, RibbonPropertyChangedEventArgs e)
        {
            var curButton = axisCombo.Current as RibbonButton;
            if (curButton == null) return;
            TsoddHost.Current.currentAxis = TsoddHost.Current.axis.FirstOrDefault(a => a.Name == curButton.Text);
        }


        // обработчик события выбора текущей стойки
        private void SplitStands_CurrentChanged(object sender, RibbonPropertyChangedEventArgs e)
        {
            var curButton = splitStands.Current as RibbonButton;
            if (curButton == null) return;
            TsoddHost.Current.currentStandBlock = curButton.Text;
        }
       
        // обработчик события выбора текущей группы знаков
        private void SignsGroups_CurrentChanged(object sender, RibbonPropertyChangedEventArgs e)
        {
            var curButton = signsGroups.Current as RibbonButton;
            if (curButton == null) return;
            TsoddHost.Current.currentSignGroup = curButton.Text;

            // пересобираем список
            splitSigns.Items.Clear();
            FillBlocksMenu(splitSigns, "SIGN", TsoddHost.Current.currentSignGroup);

            if (splitSigns.Items.Count > 0)
            {
                splitSigns.Current = splitSigns.Items[0];
            }
            else
            {
                splitSigns.Current = null;
            }

        }

        // обработчик события выбора текущего знака
        private void SplitSigns_CurrentChanged(object sender, RibbonPropertyChangedEventArgs e)
        {
            var curButton = splitSigns.Current as RibbonButton;
            if (curButton == null) return;
            TsoddHost.Current.currentSignBlock = curButton.Text;

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
                        //создаем экземпляр Axis
                        Axis axis = new Axis();

                        List<(int code, string value)> xDataList = AutocadXData.ReadXData(id);

                        // handle и id полилинии
                        axis.PolyHandle = new Handle(Convert.ToInt64(xDataList[1].value, 16));
                        try
                        {
                            axis.PolyID = db.GetObjectId(false, axis.PolyHandle, 0);
                            axis.AxisPoly = (Polyline)tr.GetObject(axis.PolyID, OpenMode.ForRead);
                        }
                        catch
                        {
                            doc.Editor.WriteMessage("\n * \n Ошибка. Не получилось получить Id полилинии. \n * \n");
                        }

                        // имя оси
                        axis.Name = xDataList[2].value;

                        // направление
                        axis.ReverseDirection = bool.Parse(xDataList[3].value);

                        // начальный пикет
                        Double.TryParse(xDataList[4].value, out axis.StartPK);

                        // добавляем ось в список осей
                        list.Add(axis);
                    }
                }
            }
            return list;
        }




        private void ApplyLineType(string name)
        {

            if (TsoddHost.Current.currentAxis == null)
            {
                EditorMessage("\nОшибка создания разметки. Не определена текущая ось \n");
                return;
            }

            if (name == null) return;
            if (_marksLineTypeFlag == false) return;

            _marksLineTypeFlag = false;

            splitMarksLineTypes.IsEnabled = false;
            rowLineType.IsEnabled = false;

            // проверяем изменения
            CheckSplitMarksLineTypesCurrentChanged(name);

            // парсим все типы линий
            var lineTypeList = LineTypeReader.Parse();

            // текущий тип линии
            var currentIAcadDef = lineTypeList.FirstOrDefault(lt => lt.Name == name);

            if (currentIAcadDef == null) { _marksLineTypeFlag = true; splitMarksLineTypes.IsEnabled = false; rowLineType.IsEnabled= true;  return; }

            // экземпляры первой и второй линии
            AcadLineType firstLineType = null;
            AcadLineType secondLineType= null;

            // создаем экземпляры текущего типа линии currentLineType_1 и ивторого типа линии currentLineType_2 (для мультилинии)
            CurrentLineType currentLineType_1 = new CurrentLineType();
            CurrentLineType currentLineType_2 = new CurrentLineType();

            // расстояние между линиями в мультилинии
            double offsetLines = 0;

            if (currentIAcadDef is AcadLineType alt) firstLineType = alt;
            if (currentIAcadDef is AcadMLineType amlt)
            {
                firstLineType = amlt.MLineLineTypes[0];
                secondLineType = amlt.MLineLineTypes[1];

                var offsetLabel = comboLineTypeOffset.Current as RibbonLabel;
                double.TryParse(offsetLabel.Description, out double val);

                // расстояние между линиями
                if (val > 0) offsetLines = val;
            }

            // заполняем первую линию
            // паттерн
            var patternLabel_1 = comboLineTypePattern_1.Current as RibbonLabel;
            var lineTypePattern_1 = patternLabel_1.Description;
            // имя типа линии в БД Autocad
            currentLineType_1.Name = $"{firstLineType.Name} {lineTypePattern_1}";
            // толщина 
            var widthLabel_1 = comboLineTypeWidth_1.Current as RibbonLabel;
            Double.TryParse(widthLabel_1.Description, out double val1_1);
            if (val1_1 >0) currentLineType_1.Width = val1_1;
            // цвет
            var colorLabel_1 = comboLineTypeColor_1.Current as RibbonLabel;
            short.TryParse(colorLabel_1.Description, out short val1_2);
            if (val1_2 > 0) currentLineType_1.ColorIndex = val1_2;

            // заполняем вторую линию для мультилинии
            if (secondLineType != null)
            {             
                // паттерн
                var patternLabel_2 = comboLineTypePattern_2.Current as RibbonLabel;
                var lineTypePattern_2 = patternLabel_2.Description;
                // имя типа линии в БД Autocad
                currentLineType_2.Name = $"{secondLineType.Name} {lineTypePattern_2}";
                // толщина 
                var widthLabel_2 = comboLineTypeWidth_2.Current as RibbonLabel;
                Double.TryParse(widthLabel_2.Description, out double val2_1);
                if (val2_1 > 0) currentLineType_2.Width = val2_1;
                // цвет
                var colorLabel_2 = comboLineTypeColor_2.Current as RibbonLabel;
                short.TryParse(colorLabel_2.Description, out short val2_2);
                if (val2_2 > 0) currentLineType_2.ColorIndex = val2_2;
            }

            // делаем выбор объектов и получаем их ID
            var listObjectsID = GetAutoCadSelectionObjectsId(new List<string> { "LWPOLYLINE", "POLYLINE" });

            if (listObjectsID == null) { _marksLineTypeFlag = true; splitMarksLineTypes.IsEnabled = true; rowLineType.IsEnabled = true; return; } // не выбрано ни одного объекта - нечего менять, выходим

            foreach (var objectId in listObjectsID)
            {
                // приводим полилинию к простой линии
                MLineTypeToPolyline(objectId, out var objId);

                if (objId == ObjectId.Null) continue;

                // если это простой тип линии (одиночная)
                if (currentIAcadDef.Type is AcadDefType.LineType)
                {
                    if (ApplyLineTypeToEntity(objId, currentLineType_1))  // проверяем полилинию и применяем оформление для нее
                    {
                        string mTextHandle = AddMtextToLineType(objId, name);

                        // создаем список с данными для XData
                        List<(int, string)> xDataList = new List<(int, string)>();
                        xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, "master"));                                              // идентификатор линии, как основной
                        xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, name));                                                  // имя типа линии                                                                                                                                      
                        xDataList.Add(((int)DxfCode.ExtendedDataHandle, TsoddHost.Current.currentAxis.PolyHandle.ToString()));        // handle оси
                        xDataList.Add(((int)DxfCode.ExtendedDataHandle, ""));                                                         // handle второй линии
                        xDataList.Add(((int)DxfCode.ExtendedDataHandle, mTextHandle));                                                // handle текста

                        AutocadXData.UpdateXData(objId, xDataList);
                    }
                }
                else    // если объект мультилиния
                {
                    Polyline masterPolyline = null;
                    Polyline slavePolyline = null;

                    PolylineTo2Polylines(objId, offsetLines, ref masterPolyline, ref slavePolyline);

                    // если получилось преобразовать в полилинию
                    if (masterPolyline != null && slavePolyline != null)
                    {
                        // проверяем второстепенную полилинию и применяем оформление для нее
                        if (ApplyLineTypeToEntity(slavePolyline.Id, currentLineType_2))
                        {
                            // создаем список с данными для XData
                            List<(int, string)> xDataList = new List<(int, string)>();
                            xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, "slave"));                                             // идентификатор линии, как ведомой 
                            xDataList.Add(((int)DxfCode.ExtendedDataHandle, masterPolyline.Handle.ToString()));                         // handle основной линии
                      
                            AutocadXData.UpdateXData(slavePolyline.Id, xDataList);
                        }

                        // проверяем основную полилинию и применяем оформление для нее
                        if (ApplyLineTypeToEntity(masterPolyline.Id, currentLineType_1))  
                        {
                            string mTextHandle = AddMtextToLineType(masterPolyline.Id, name);

                            // создаем список с данными для XData
                            List<(int, string)> xDataList = new List<(int, string)>();
                            xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, "master"));                                              // идентификатор линии, как основной 
                            xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, name));                                                  // имя типа линии
                            xDataList.Add(((int)DxfCode.ExtendedDataHandle, TsoddHost.Current.currentAxis.PolyHandle.ToString()));        // handle оси
                            xDataList.Add(((int)DxfCode.ExtendedDataHandle, slavePolyline.Handle.ToString()));                            // handle второй линии
                            xDataList.Add(((int)DxfCode.ExtendedDataHandle, mTextHandle));                                                // handle текста


                            AutocadXData.UpdateXData(masterPolyline.Id, xDataList);
                        }

                    }
                }
            }

            _marksLineTypeFlag = true;
            splitMarksLineTypes.IsEnabled = true;
            rowLineType.IsEnabled = true;
        }




        private void CheckSplitMarksLineTypesCurrentChanged(string name)
        {
            // var curButton = splitMarksLineTypes.Current as RibbonButton;
            // if (curButton == null) return;

            if(TsoddHost.Current.currentMarksLineType == name) return;  // это не изменение типа линии, а просто применение

            TsoddHost.Current.currentMarksLineType = name;  // запоминаем имя текущего типа линии

            // парсим все типы линий
            var lineTypeList = LineTypeReader.Parse();

            var currentLineType = lineTypeList.FirstOrDefault(lt => lt.Name == name);
            if (currentLineType == null) return;

            // очищаем все
            comboLineTypePattern_1.Items.Clear();
            comboLineTypePattern_2.Items.Clear();
            comboLineTypeWidth_1.Items.Clear();
            comboLineTypeWidth_2.Items.Clear();
            comboLineTypeColor_1.Items.Clear();
            comboLineTypeColor_2.Items.Clear();
            comboLineTypeOffset.Items.Clear();

            switch (currentLineType)
            {
                case var l when l is AcadLineType lt:   // когда это простой тип линии
                    
                    // блокируем контролы для второй линии
                    comboLineTypePattern_2.IsVisible = false;
                    comboLineTypeWidth_2.IsVisible = false;
                    comboLineTypeColor_2.IsVisible = false;
                    comboLineTypeOffset.IsVisible = false;
                    labelLineType.IsVisible = false;

                    // заполняем паттерны
                    FillPattern(lt.PatternValues, comboLineTypePattern_1);
                    // заполняем толщины
                    FillWidth(lt.Width, comboLineTypeWidth_1);
                    // заполняем цвета
                    FillColor(lt.ColorIndex, comboLineTypeColor_1);

                    break;

                case var l when l is AcadMLineType mlt:

                    // разблокируем контролы для второй линии
                    comboLineTypePattern_2.IsVisible = true;
                    comboLineTypeWidth_2.IsVisible = true;
                    comboLineTypeColor_2.IsVisible = true;
                    comboLineTypeOffset.IsVisible = true;
                    labelLineType.IsVisible = true;


                    if (mlt.MLineLineTypes.Count == 2)      // может быть только два элемента
                    {
                        var first_line = mlt.MLineLineTypes[0];
                        if (first_line != null)
                        {
                            //заполняем паттерны
                            FillPattern(first_line.PatternValues, comboLineTypePattern_1);
                            // заполняем толщины
                            FillWidth(first_line.Width, comboLineTypeWidth_1);
                            // заполняем цвета
                            FillColor(first_line.ColorIndex, comboLineTypeColor_1);
                        }

                        var second_line = mlt.MLineLineTypes[1];
                        if (second_line != null)
                        {
                            //заполняем паттерны
                            FillPattern(second_line.PatternValues, comboLineTypePattern_2);
                            // заполняем толщины
                            FillWidth(second_line.Width, comboLineTypeWidth_2);
                            // заполняем цвета
                            FillColor(second_line.ColorIndex, comboLineTypeColor_2);
                        }

                        // заполняем расстояния между линиями
                        FillOffset(mlt.Offset, comboLineTypeOffset);
                    }
                    break;

            }

            // внутренний метод заполнения combobox паттерна
            void FillPattern(List<List<double>> list, RibbonCombo combo)
            {
                foreach (var pattern in list)
                {
                    string txt = "";
                    string descriprionTxt = "";

                    foreach (var val in pattern)
                    {
                        descriprionTxt += $"_{val}";
                        if (val == 1000)
                        {
                            txt = "сплошная";
                            descriprionTxt = "_continuous";
                            break;
                        }
                        txt += $" [{val}] ";
                    }
                    if (txt != "")
                    {
                        descriprionTxt = descriprionTxt.Substring(1).Trim();
                        RibbonLabel rbl = new RibbonLabel { Text = txt, Description = $"({descriprionTxt})" };
                        combo.Items.Add(rbl);
                    }
                }
                if (combo.Items.Count > 0) combo.Current = combo.Items[0];  // выбираем первый
            }


            // внутренний метод заполнения combobox width
            void FillWidth(List<double> list, RibbonCombo combo)
            {
                foreach (var width in list)
                {
                    RibbonLabel rbl = new RibbonLabel { Text = width.ToString(), Description = width.ToString() };
                    combo.Items.Add(rbl);
                }
                if (combo.Items.Count > 0) combo.Current = combo.Items[0];  // выбираем первый
            }


            // внутренний метод заполнения combobox color
            void FillColor(List<short>list, RibbonCombo combo)
            {
                foreach (var color in list)
                {
                    var bmp = new Bitmap(12, 12);
                    using (var g = Graphics.FromImage(bmp))     // холст созданный, по Bitmap
                    {
                        Autodesk.AutoCAD.Colors.Color cl = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, (short)color);

                        using (var pen = new Pen(cl.ColorValue, 1f))  //  Кисть для рисовния 
                        {
                            // прозрачный фон bmp 
                            g.Clear(Color.Transparent);

                            // прямоугольник
                            var rect = new Rectangle(0, 0, 12, 12);
                    
                            g.DrawRectangle(pen, rect);
                            using (var brush = new SolidBrush(cl.ColorValue))
                            {
                                g.FillRectangle(new SolidBrush(cl.ColorValue), rect);
                            }

                            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                            var bmpSource = TsoddBlock.ToImageSource(bmp);

                            RibbonLabel rbl = new RibbonLabel { Image = bmpSource, LargeImage = bmpSource, Description = $"{color}" };
                            combo.Items.Add(rbl);   
                        }
                    } 
                }
                if (combo.Items.Count > 0) combo.Current = combo.Items[0];  // выбираем первый
            }

            //внутренний метод заполнения combobox расстояния между линиями
            void FillOffset(List<double> list, RibbonCombo combo)
            {
                foreach (var offset in list)
                {
                    RibbonLabel rbl = new RibbonLabel { Text = offset.ToString(), Description = offset.ToString() };
                    combo.Items.Add(rbl);
                }
                if (combo.Items.Count > 0) combo.Current = combo.Items[0];  // выбираем первый
            }
        }



        /// <summary>
        /// Метод убирает из списка все элементы slave
        /// </summary>
        /// <param name="listId" список ObjectId элементов></param>
        private void DeleteSlaveObjectsFromList(ref List<ObjectId> listId)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (var id in listId)
                    {
                        // получаем xData объекта
                        var xDataList = AutocadXData.ReadXData(id);

                        if (xDataList.Count == 0) continue;

                        // проверяем наличие slave
                        if (xDataList.Any(d => d.Item2 == "slave"))
                        {
                            listId.Remove(id);  // удаляем из списка элемент
                        }
                    }
                }
            }
        }





        /// <summary>
        /// Метод преобразавывает полилинию (переданный id) в две полилинии на расстоянии offset от друг-друга
        /// </summary>
        /// <param name="id" ID идентификатор выбранной полилинии></param>
        /// <param name="offset" расстояние между результирующими полилиниями></param>
        /// <param name="polyline_1" результирующая полилиния 1 ></param>
        /// <param name="polyline_2" результирующая полилиния 2 ></param>
        private void PolylineTo2Polylines(ObjectId id, double offset, ref Polyline polyline_1, ref Polyline polyline_2)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // получаем полилинию по Id
                    var polyline = (Polyline)tr.GetObject(id, OpenMode.ForRead);

                    // проверка на то, что полилиния не является осью
                    var axis = TsoddHost.Current.axis.FirstOrDefault(a => a.PolyHandle == polyline.Handle);
                    if (axis != null)
                    {
                        MessageBox.Show($" Ошибка создания разметки. Данная полилиния уже была выбрана для оси \" {axis.Name} \" ");
                        tr.Abort();
                        return;
                    }

                    // считаем offset
                    offset = db.Insunits == UnitsValue.Millimeters ? offset : offset*1000;

                    // пробуем получить подобные полилинии
                    try { polyline_1 = (Polyline)polyline.GetOffsetCurves(offset/2)[0]; }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex){ ed.WriteMessage($"Не получилось создать подобную линию для полилинии с Id = {id}"); }

                    try { polyline_2 = (Polyline)polyline.GetOffsetCurves(-offset/2)[0]; }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex) { ed.WriteMessage($"Не получилось создать подобную линию для полилинии с Id = {id}"); }

                    // добавляем линии в чертеж
                    if (polyline_1 != null && polyline_2 != null) 
                    {
                      
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        btr.AppendEntity(polyline_1);
                        tr.AddNewlyCreatedDBObject(polyline_1, true);

                        btr.AppendEntity(polyline_2);
                        tr.AddNewlyCreatedDBObject(polyline_2, true);

                        // копируем xData исходного элемента в полилинию 1
                        var temListXdata = AutocadXData.ReadXData(polyline.Id);
                        AutocadXData.UpdateXData(polyline_1.Id, temListXdata);

                        // удаляем исходный элемкнт
                        var ent = (Entity)tr.GetObject(id, OpenMode.ForWrite, openErased: true);
                        if (ent != null && !ent.IsErased)
                        {
                            AutocadXData.UpdateXData(id, null);
                            ent.Erase(true);
                        }

                        tr.Commit();

                    }
                }
            }
        }


        /// <summary>
        /// Метод проверяет, является ли объект MlineType и приводит его к простой полилинии
        /// </summary>
        /// <param name="id" ID идентификатор выбранной полилинии></param>
        public static void MLineTypeToPolyline(ObjectId id, out ObjectId outId)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            outId = id;


            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {

                    // проверка на то, что объект еще есть в бд
                    try { var dbo = (DBObject)tr.GetObject(id, OpenMode.ForRead); }
                    catch { outId = ObjectId.Null ; return;  }  // если объекта нет, то выходим

                    TsoddXdataElement tsoddXdataElement = new TsoddXdataElement();
                    tsoddXdataElement.Parse(id);

                    // если  есть и master и slave полилиния, то понимаем, что объект был полилиние, а значит пора преобразовать в одну полилинию
                    if (tsoddXdataElement.MasterPolylineID != ObjectId.Null)
                    {
                        // удаляем старый текст
                        if (tsoddXdataElement.MtextID!= ObjectId.Null)
                        {
                            // удаляем старый mtext
                            var dbo = (DBObject)tr.GetObject(tsoddXdataElement.MtextID, OpenMode.ForWrite, openErased: true);
                            if (dbo != null && !dbo.IsErased)
                            {
                                dbo.Erase(true);
                            }
                        }

                        if (tsoddXdataElement.SlavePolylineID != ObjectId.Null)
                        {
                            Polyline masterPolyline = (Polyline)tr.GetObject(tsoddXdataElement.MasterPolylineID, OpenMode.ForRead);        // master полилиния
                            Polyline slavePolyline = (Polyline)tr.GetObject(tsoddXdataElement.SlavePolylineID, OpenMode.ForRead);        // master полилиния

                            // расстояние между линиями 
                            double dist = masterPolyline.StartPoint.DistanceTo(slavePolyline.StartPoint);

                            // оффнтим masterPolyline на половину полученного расстояния и проверяем сторону оффсета
                            Polyline tempPoliline_1 = null;
                            Polyline tempPoliline_2 = null;

                            try { tempPoliline_1 = (Polyline)masterPolyline.GetOffsetCurves(dist / 2)[0]; }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex) { ed.WriteMessage($"Не получилось создать подобную линию для полилинии с Id = {id}"); }

                            try { tempPoliline_2 = (Polyline)masterPolyline.GetOffsetCurves(-dist / 2)[0]; }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex) { ed.WriteMessage($"Не получилось создать подобную линию для полилинии с Id = {id}"); }

                            // проверка расстояния
                            double dist_1 = tempPoliline_1.StartPoint.DistanceTo(slavePolyline.StartPoint);
                            double dist_2 = tempPoliline_2.StartPoint.DistanceTo(slavePolyline.StartPoint);

                            Polyline resultPolyline = dist_1 < dist_2 ? tempPoliline_1 : tempPoliline_2;

                            // добавляем результирующую полилинию в чертеж
                            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                            btr.AppendEntity(resultPolyline);
                            tr.AddNewlyCreatedDBObject(resultPolyline, true);

                            //меняем ObjectId на новую линию
                            outId = resultPolyline.Id;

                            // копируем xData исходного элемента в полилинию 1
                            var temListXdata = AutocadXData.ReadXData(masterPolyline.Id);
                            AutocadXData.UpdateXData(resultPolyline.Id, temListXdata);

                            // удаляем старые полилинии
                            var dbo_1 = (DBObject)tr.GetObject(masterPolyline.Id, OpenMode.ForWrite);
                            if (dbo_1 != null && !dbo_1.IsErased)
                            {
                                AutocadXData.UpdateXData(masterPolyline.Id, null);
                                dbo_1.Erase(true);
                            }
                            var dbo_2 = (DBObject)tr.GetObject(slavePolyline.Id, OpenMode.ForWrite);
                            if (dbo_2 != null && !dbo_2.IsErased)
                            {
                                AutocadXData.UpdateXData(slavePolyline.Id, null);
                                dbo_2.Erase(true);
                            }
                        }

                        tr.Commit();
                    }
                }
            }
        }




        //// тестовый метод
        //public static void Test()
        //{
        //    var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        //    var db = doc.Database;
        //    var ed = doc.Editor;

        //    // получаем список выбранных элементов
        //    var listElementsID = RibbonInitializer.Instance.GetAutoCadSelectionObjectsId(new List<string> { "LWPOLYLINE" });

        //    using (doc.LockDocument())
        //    {
        //        using (var tr = db.TransactionManager.StartTransaction())
        //        {
        //            // проходимся по всем элементам выбора
        //            foreach (var id in listElementsID)
        //            {
        //                var entity = (Entity)tr.GetObject(id, OpenMode.ForRead);
        //                if (entity.XData != null)   // если у объекта есть xData
        //                {
        //                    // получаем XData исходного объекта
        //                    List<(int code, string value)> xDataList = AutocadXData.ReadXData(id);

        //                    if (xDataList.Count == 0) continue;  // если нет xData то это не объект MlineType

        //                    Polyline masterPolyline = null;
        //                    Polyline slavePolyline = null;
        //                    ObjectId masterPolylineId = ObjectId.Null;
        //                    ObjectId slavePolylineId = ObjectId.Null;

        //                    switch (xDataList[0].Item2)
        //                    {
        //                        case "slave":
        //                            var masterPolylineHandle = new Handle(Convert.ToInt64(xDataList[1].value, 16));     // handle master полилинии
        //                            masterPolylineId = db.GetObjectId(false, masterPolylineHandle, 0);              // Id master полилинии
        //                            masterPolyline = (Polyline)tr.GetObject(masterPolylineId, OpenMode.ForRead);        // master полилиния

        //                            slavePolyline = (Polyline)tr.GetObject(id, OpenMode.ForRead);                       // slave полилиния

        //                            break;

        //                        case "master":
        //                            masterPolyline = (Polyline)tr.GetObject(id, OpenMode.ForRead);                       // master полилиния
        //                            masterPolylineId = id;

        //                            var slavePolylineHandle = new Handle(Convert.ToInt64(xDataList[3].value, 16));       // handle slave полилинии

        //                            if (!string.Equals(slavePolylineHandle.ToString(), "0"))                 // это не мултилиня, дальнейшие действия не нужны
        //                            {
        //                                slavePolylineId = db.GetObjectId(false, slavePolylineHandle, 0);                     // Id master полилинии
        //                                slavePolyline = (Polyline)tr.GetObject(slavePolylineId, OpenMode.ForRead);           // slave полилиния
        //                            }
        //                            break;

        //                        default: return;
        //                    }


        //                }

        //            }

        //        }
        //    }
        //}



        /// <summary>
        /// Метод применяет выбранное оформление типа линии для полилинии (разметка)
        /// </summary>
        /// <param name="id">ID идентификатор выбранной полилинии.</param>
        /// <param name="lineType">Параметры оформления полилинии.</param>
        /// <returns> bool - удачное применение оформления или нет.</returns>
        private bool ApplyLineTypeToEntity(ObjectId id, CurrentLineType lineType)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {

                    var polyline = (Polyline)tr.GetObject(id, OpenMode.ForWrite);
 
                    // проверка на то, что полилиния не является осью
                    var axis = TsoddHost.Current.axis.FirstOrDefault(a => a.PolyHandle == polyline.Handle);
                    if (axis != null)
                    {
                        MessageBox.Show($" Ошибка создания разметки. Данная полилиния уже была выбрана для оси \" {axis.Name} \" ");
                        tr.Abort();
                        return false;
                    }

                    polyline.Linetype = lineType.Name;
                    polyline.ConstantWidth = lineType.Width;
                    polyline.ColorIndex = lineType.ColorIndex;

                    tr.Commit();
                    return true;
                }
            }
        }


        /// <summary>
        /// Метод добавляет описание типа линии на чертеж (номер типа линии и длину полилинии в метрах)
        /// </summary>
        /// <param name="polylineID"> ID идентификатор выбранной полилинии. </param>
        /// <param name="lineTypeName"> Наименование типа линии</param>
        /// <returns> Возвращает handle, созданного MText </returns>
        private string AddMtextToLineType(ObjectId polylineID, string lineTypeName)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            if (polylineID == null) return null;

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
   
                    // исходная полилиния
                    Polyline polyline = (Polyline)tr.GetObject(polylineID, OpenMode.ForRead);
                   
                    // уточняем положение текста
                    Polyline tempPoly = null;
                    double scale = db.Cannoscale.DrawingUnits;
                    double scaleFactor = db.Insunits == UnitsValue.Millimeters ? 0.001 : 1;

                    try { tempPoly = (Polyline)polyline.GetOffsetCurves(2.5 * scale)[0]; }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex) {}

                    if (tempPoly == null) tempPoly = polyline;  // лесли не получилось создать оффсет полилинии, то возьмем текущую

                    // проверим является ли средний сегмент дугой в полилинии
                    Point3d centerPoint = tempPoly.GetPointAtDist(tempPoly.Length / 2);
                    double parametr = tempPoly.GetParameterAtPoint(centerPoint);
                    Vector3d tan = tempPoly.GetFirstDerivative(parametr).GetNormal();
                   
                    double angle = Math.Atan2(tan.Y, tan.X);
                    
                    var handle = polylineID.Handle.ToString();
                    string mTextId = polylineID.OldId.ToString();

                    MText mt = new MText();
                    mt.Annotative = AnnotativeStates.True;
                    mt.Location = centerPoint;
                    mt.Rotation = angle;
                    mt.Contents = $"{lineTypeName} ("+ $@"%<\AcObjProp Object(%<\_ObjId {mTextId}>%).Length \f ""%lu2%ct8[{scaleFactor}]"">%" + " м)";
                    mt.TextHeight = 2.5 * scale;
                    mt.Height = 2.5;
                    mt.Attachment = AttachmentPoint.MiddleCenter;

                    btr.AppendEntity(mt);
                    tr.AddNewlyCreatedDBObject(mt, true);

                    tr.Commit();
                    ed.Regen();

                    // записть XData в мтекст
                    List<(int, string)> mTextXdataList = new List<(int, string)>();
                    mTextXdataList.Add(((int)DxfCode.ExtendedDataAsciiString, "slave"));                                               // идентификатор  - ведомый элемент
                    mTextXdataList.Add(((int)DxfCode.ExtendedDataHandle, handle));                                                     // handle основной линии

                    AutocadXData.UpdateXData(mt.Id, mTextXdataList);

                    return mt.Handle.ToString();
                }
            }


        }



        // метод перестроения и сохранения объектов TSODD
        //private void SaveTsoddChanges()
        //{
        //    // обновим блоки осей
        //    foreach (Axis axis in TsoddHost.Current.axis)
        //    {
        //        TsoddBlock.AddAxisBlock(axis);
        //    }
        //}



        // метод создает список RibbonButton для каждого кастомного типа линий
        private void ListOfMarksLinesLoad(int bmpWidth, int bmpHeight, string preSelect = null)
        {
            splitMarksLineTypes.Items.Clear();

            // проходимся по всем типам линий в файле
            foreach (var lineType in LineTypeReader.Parse())
            {
                if (lineType.IsMlineElement)  continue;  //  если это составная часть мультилинии, то пропускаем 

                var bmp = new Bitmap(bmpWidth, bmpHeight); // новый Bitmap для отрисовки типа линии

                using (var g = Graphics.FromImage(bmp))     // холст созданный, по Bitmap
                {
                    using (var pen = new Pen(Color.WhiteSmoke, 2f))  //  Кисть для рисовния 
                    {
                        // прозрачный фон bmp 
                        g.Clear(Color.Transparent);

                        // подпись слоя
                        using (var font = new System.Drawing.Font("Mipgost", 10f, System.Drawing.FontStyle.Bold ))
                        {
                            g.DrawString(lineType.Name, font, Brushes.WhiteSmoke, new PointF(0, 2));
                        }

                        // получаем паттерн штриховой линии простого типа лини
                        if (lineType is AcadLineType lt && lt.PatternValues.Count > 0)
                        {
                            float[] patternArray = GetPatternArray(lt.PatternValues[0],5);
                            pen.DashPattern = patternArray;
                            int y = bmpHeight / 2+1;
                            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                            g.DrawLine(pen, 50, y, bmpWidth - 6, y);
                        }

                        // если это мультилиния, то находим все линии и берем паттерн штриховки у них
                        if (lineType is AcadMLineType mlt)
                        {
                            int y = bmpHeight;
                            foreach (var line in mlt.MLineLineTypes)
                            {
                                float[] patternArray = GetPatternArray(line.PatternValues[0], 5);
                                pen.DashPattern = patternArray;
                                y -= bmpHeight / 3+1;
                                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                                g.DrawLine(pen, 50, y, bmpWidth - 6, y);
                            }
                        }

                        var bmpSource = TsoddBlock.ToImageSource(bmp);
                        RibbonButton rb = new RibbonButton()
                        {
                            LargeImage = bmpSource,
                            Image = bmpSource,
                            ShowText = false,
                            Description = lineType.Name,
                            ShowImage = true,
              
                        };

                        rb.CommandHandler = new RelayCommandHandler(() =>
                        {
                            ApplyLineType((string)rb.Description);
                        });

                        splitMarksLineTypes.Items.Add(rb);
                    }
                }
            }

            // предвыбор
            if (splitMarksLineTypes.Items.Count > 0 && preSelect == null)
            {
                if (preSelect != null)
                {
                    var item = splitMarksLineTypes.Items.FirstOrDefault(lt => lt.Description == preSelect);
                    if (item != null)
                    {
                        splitMarksLineTypes.Current = item;
                        CheckSplitMarksLineTypesCurrentChanged((string)item.Description);
                    }
                }
                else
                {
                    splitMarksLineTypes.Current = splitMarksLineTypes.Items[0];
                    CheckSplitMarksLineTypesCurrentChanged((string)splitMarksLineTypes.Items[0].Description);
                }
            }

        }



        // метод трансформации паттерна штриховки для линий
        private float[] GetPatternArray(List<double> lineTypeValues, float scale)
        {
            float[] patternArray = null;
            List<float> patternList = new List<float>();

            for (int i = 0; i < lineTypeValues.Count; i++)
            {
                // текущее значенеие
                float cur_val = (float)lineTypeValues[i];
                float last_val = 0;

                // сравнение с предыдущим значением
                if (i > 0)
                {
                    last_val = (float)lineTypeValues[i - 1];

                    // если текущее или предыдущее значение было точкой то сразу записываем
                    if (cur_val == 0 || last_val == 0)
                    {
                        float val = cur_val == 0 ? 1 : cur_val * scale;

                        patternList.Add(Math.Abs(val));
                        continue;
                    }
                    else
                    {
                        // если текущее и предыдущее значения оба положительные или отрицательные
                        if ((cur_val > 0 && last_val > 0) || (cur_val < 0 && last_val < 0))
                        {
                            // обновляем последнее значение
                            patternList[patternList.Count - 1] = patternList.Last() + cur_val * scale;
                            continue;
                        }

                        // если ни одна проверка не сработала, то просто добавим это значение
                        patternList.Add(Math.Abs(cur_val * scale));
                    }
                }
                else    // если это первый элемент, то просто его добавим
                {
                    patternList.Add(Math.Abs(cur_val * scale));
                }

            }

           patternArray = patternList.ToArray();
           //pen.DashPattern = patternArray;

           return patternArray;
        }

        // заполняет RibbonSplitButton элементами RibbonButton с именами и картинками
        private void FillBlocksMenu(RibbonSplitButton split, string tag,  string preselect = null)
        {
            split.Items.Clear();

            //если это знаки, то будем получать список 
            string group = null;
            if (tag == "SIGN" && TsoddHost.Current.currentSignGroup != null) group = TsoddHost.Current.currentSignGroup;

            var listOfBlocks =TsoddBlock.GetListOfBlocks(tag,group);    // формируем список блоков по тегу

            foreach (var block in listOfBlocks)                         // для всех блков сосздаем кнопку RibbonButton с имененм и картинкой
            {
                var item = new RibbonButton
                {
                    Text = block.name,
                    ShowText = true,
                    ShowImage = block.img != null,
                    Image = block.img,      // иконка в меню
                    LargeImage = block.img, // пусть будет и large
                    Size = RibbonItemSize.Standard,
                    Orientation = System.Windows.Controls.Orientation.Horizontal,
                    CommandParameter = block.name
                };

                item.CommandHandler = new RelayCommandHandler(() =>
                {
                    var name = (string)item.CommandParameter;
                    switch (tag)
                    {
                        case "STAND":
                            TsoddBlock.InsertStandBlock(name);
                            break;
                        case "SIGN":
                            TsoddBlock.InsertSignBlock(name);
                            break;
                    }
 
                });
                split.Items.Add(item);
            }

            // Если нужен предвыбор
            if (preselect != null)
            {
                var ribButton = split.Items.OfType<RibbonButton>().FirstOrDefault(r => r.Text == preselect);
                if (ribButton != null) split.Current = ribButton;
            }
        }


        // метод выбора элементов в AutoCAD с фильтром
        public List<ObjectId> GetAutoCadSelectionObjectsId (List<string>filterList)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            filterList = filterList.Select(a => a.ToUpper()).ToList(); // все в верхний регистр

            // массив TypedValue[] для фильтра
            TypedValue[] typedValue = new TypedValue[filterList.Count+2];
            typedValue[0] = new TypedValue((int)DxfCode.Operator, "<OR");
            typedValue[typedValue.Length-1] = new TypedValue((int)DxfCode.Operator, "OR>");
            for (int i = 1; i < typedValue.Length-1; i++) typedValue[i] = new TypedValue((int)DxfCode.Start, filterList[i-1]);

            var filter = new SelectionFilter(typedValue);

            var pso = new PromptSelectionOptions
            {
                MessageForAdding = "\n Выберите объекты, можно рамкой выбрать несколько:" ,
                SingleOnly = false , // можно выбрать несколько
                AllowDuplicates = false
            };

            var psr = ed.GetSelection(pso, filter);

            if (psr.Status != PromptStatus.OK) return null; // неудачный промпт, выходим

            List<ObjectId> resultList = new List<ObjectId>();
            foreach (var id in psr.Value.GetObjectIds()) resultList.Add(id);

            return resultList;
        }


        









        // метод подгружающий картинки для кнопок
        private BitmapImage LoadImage(string uri)
        {
            return new BitmapImage(new Uri(uri, UriKind.Absolute));

        }

                                                    
        // вывод сообщения в editor
        private void EditorMessage(string txt)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            ed.WriteMessage($" \n {txt}");
        }








    }
}
