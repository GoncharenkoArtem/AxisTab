using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows;
using System.Windows.Media.Imaging;
using TSODD;
using TSODD.Forms;


/* Методы и обработчики для Ribbon */
namespace TSODD
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

                e.Document.Editor.PromptForSelectionEnding += Editor_PromptForSelectionEnding;
            }
            ListOFAxisRebuild(); // перестраиваем combobox с осями
            LineTypeReader.RefreshLineTypesInAcad(); // загружаем типы линии
            TsoddBlock.blockInsertFlag = true; // разрешаем вставку блоков
        }



        private void Editor_PromptForSelectionEnding(object sender, PromptForSelectionEndingEventArgs e)
        {
            if (!quickPropertiesOn) return; // если отключено окно с бысрыми свойствами

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // Id всех выбранных объектов и которые являются вставками блока
            var selectedIds = e.Selection.Cast<SelectedObject>()
                    .Select(o => o.ObjectId)
                    .Where(id => id.ObjectClass.Name == RXClass.GetClass(typeof(BlockReference)).Name ||
                                 id.ObjectClass.Name == RXClass.GetClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline)).Name)
                    .ToList();


            if (selectedIds.Count == 0)
            {
                if (selectioForm != null) { selectioForm.Close(); selectioForm = null; }
                return; // если нет подходящих объектов, то выходим 
            }

            // получаем словарь с нужными блоками и типами линий (SIGN,MARK)
            Dictionary<ObjectId, string> dictionary = GetDictionaryOfCorrectBlocks(selectedIds);

            if (dictionary.Count == 0)
            {
                if (selectioForm != null) { selectioForm.Close(); selectioForm = null; }
                return; // если нет подходящих объектов, то выходим 
            }

            // Вычисляем тип формы для выбора блоков 
            bool signs = dictionary.Values.Any(m => m == "SIGN");
            bool marks = dictionary.Values.Any(m => m == "MARK");

            int formType = 0;
            switch (signs, marks)
            {
                case var type when type.signs == true && type.marks == false:
                    formType = 0;
                    break;

                case var type when type.signs == false && type.marks == true:
                    formType = 1;
                    break;

                case var type when type.signs == true && type.marks == true:
                    formType = 2;
                    break;
            }


            // проверяем надо ли вызвать форму или перестроить ее 
            if (selectioForm == null)
            {
                selectioForm = new SelectionFormBlocks(formType, dictionary);
                selectioForm.Show();
                selectioForm.Focus();
            }
            else
            {
                // проверяем надо ли перерисовать форму
                bool dictionariesAreEqual = selectioForm._dictionary.Count == dictionary.Count && selectioForm._dictionary.OrderBy(k => k.Key).SequenceEqual(dictionary.OrderBy(k => k.Key));
                if (selectioForm._type != formType || !dictionariesAreEqual) selectioForm.RebuildForm(formType, dictionary);
                selectioForm.Focus();
            }


            // внутренний метод для заполния словаря
            Dictionary<ObjectId, string> GetDictionaryOfCorrectBlocks(List<ObjectId> ids)
            {
                Dictionary<ObjectId, string> blockDictionary = new Dictionary<ObjectId, string>();

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // проходимся по всем объектам и ищем знаки и разметку
                    foreach (ObjectId id in ids)
                    {
                        DBObject dbo = (DBObject)tr.GetObject(id, OpenMode.ForRead);
                        if (dbo is BlockReference br)   // если это блок
                        {
                            foreach (ObjectId attr in br.AttributeCollection)
                            {
                                var at = tr.GetObject(attr, OpenMode.ForRead) as AttributeReference;
                                if (at.Tag.Equals("SIGN", StringComparison.OrdinalIgnoreCase)) { blockDictionary[id] = "SIGN"; break; }
                                if (at.Tag.Equals("MARK", StringComparison.OrdinalIgnoreCase)) { blockDictionary[id] = "MARK"; break; }
                            }
                        }

                        if (dbo is Autodesk.AutoCAD.DatabaseServices.Polyline poly)   //  если это полилиния
                        {
                            // проверяем есть ли xData у полилинии
                            if (poly.XData != null)
                            {
                                TsoddXdataElement tsoddXdataElement = new TsoddXdataElement();
                                tsoddXdataElement.Parse(poly.Id);

                                if (tsoddXdataElement.Type == TsoddElement.Axis) continue;  // если это ось, то пропускаем объект
                                blockDictionary[tsoddXdataElement.MasterPolylineID] = "MARK";
                            }
                        }
                    }
                }
                return blockDictionary;
            }

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
                e.Document.Editor.PromptForSelectionEnding -= Editor_PromptForSelectionEnding;

                db.ObjectErased -= Db_ObjectErased;
            }

            TsoddHost.tsoddDictionary.Remove(dockey);

        }


        // обработчик удаления элемента из базы autoCAD
        private void Db_ObjectErased(object sender, Autodesk.AutoCAD.DatabaseServices.ObjectErasedEventArgs e)
        {
            if (!readyToDeleteEntity) return; // если не запрещен анализ Xdata другим процессами
            if (!e.Erased) return;
            if (e.DBObject.GetType().Name != "Polyline") return;

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;


            // проверяем есть ли xData у удаляемого объекта
            if (e.DBObject.XData != null)
            {
                TsoddXdataElement tsoddXdataElement = new TsoddXdataElement();
                try { tsoddXdataElement.Parse(e.DBObject.Id); }
                catch (Autodesk.AutoCAD.Runtime.Exception ex) { return; }

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
                else      // если это линия или мультилиния
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


        // метод выбора и проверки полилинии оси
        public Axis SelectAxis()
        {
            Axis resultAxis = null;
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            // Настройки промпта
            var peo = new PromptEntityOptions("\n Выберите ось: ");
            peo.SetRejectMessage("\n Это не полилиния. Выберите полилинию!");
            peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), exactMatch: false);

            // Запрос
            var per = ed.GetEntity(peo);
            if (per.Status == PromptStatus.OK)
            {
                resultAxis = TsoddHost.Current.axis.FirstOrDefault(a => a.PolyID == per.ObjectId);
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

            foreach (var a in TsoddHost.Current.axis)
                axisCombo.Items.Add(new RibbonButton { Text = a.Name, ShowText = true });

            // выбрать по умолчанию
            var toSelect = preselect != null
                ? axisCombo.Items.OfType<RibbonButton>().FirstOrDefault(b => b.Text == preselect.Name)
                : axisCombo.Items.OfType<RibbonButton>().FirstOrDefault();

            if (toSelect != null)
                axisCombo.Current = toSelect;

            if (preselect != null) TsoddHost.Current.currentAxis = preselect;
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
                            axis.AxisPoly = (Autodesk.AutoCAD.DatabaseServices.Polyline)tr.GetObject(axis.PolyID, OpenMode.ForRead);
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

            if (currentIAcadDef == null) { _marksLineTypeFlag = true; splitMarksLineTypes.IsEnabled = false; rowLineType.IsEnabled = true; return; }

            // экземпляры первой и второй линии
            AcadLineType firstLineType = null;
            AcadLineType secondLineType = null;

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
            if (lineTypePattern_1.Contains("CONTINUOUS")) { currentLineType_1.Name = "CONTINUOUS"; }
            else{ currentLineType_1.Name = $"{firstLineType.Name} {lineTypePattern_1}"; }    
            // толщина 
            var widthLabel_1 = comboLineTypeWidth_1.Current as RibbonLabel;
            Double.TryParse(widthLabel_1.Description, out double val1_1);
            if (val1_1 > 0) currentLineType_1.Width = val1_1;
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
                if (lineTypePattern_2.Contains("CONTINUOUS")) { currentLineType_2.Name = "CONTINUOUS"; }
                else { currentLineType_2.Name = $"{secondLineType.Name} {lineTypePattern_2}"; }
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
                        xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, "master"));                                              // 0 идентификатор линии, как основной
                        xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, name));                                                  // 1 имя типа линии                                                                                                                                      
                        xDataList.Add(((int)DxfCode.ExtendedDataHandle, TsoddHost.Current.currentAxis.PolyHandle.ToString()));        // 2 handle оси
                        xDataList.Add(((int)DxfCode.ExtendedDataHandle, ""));                                                         // 3 handle второй линии
                        xDataList.Add(((int)DxfCode.ExtendedDataHandle, mTextHandle));                                                // 4 handle текста
                        xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, "Холодный пластик"));                                    // 5 материал
                        xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, "Нанести"));                                             // 6 наличие

                        AutocadXData.UpdateXData(objId, xDataList);
                    }
                }
                else    // если объект мультилиния
                {
                    Autodesk.AutoCAD.DatabaseServices.Polyline masterPolyline = null;
                    Autodesk.AutoCAD.DatabaseServices.Polyline slavePolyline = null;

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
                            xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, "master"));                                              // 0 идентификатор линии, как основной 
                            xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, name));                                                  // 1 имя типа линии
                            xDataList.Add(((int)DxfCode.ExtendedDataHandle, TsoddHost.Current.currentAxis.PolyHandle.ToString()));        // 2 handle оси
                            xDataList.Add(((int)DxfCode.ExtendedDataHandle, slavePolyline.Handle.ToString()));                            // 3 handle второй линии
                            xDataList.Add(((int)DxfCode.ExtendedDataHandle, mTextHandle));                                                // 4 handle текста
                            xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, "Холодный пластик"));                                    // 5 материал
                            xDataList.Add(((int)DxfCode.ExtendedDataAsciiString, "Нанести"));                                             // 6 наличие

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

            if (TsoddHost.Current.currentMarksLineType == name) return;  // это не изменение типа линии, а просто применение

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
                        if (val == 1000)
                        {
                            txt = "сплошная";
                            descriprionTxt = "_CONTINUOUS";
                            break;
                        }
                        descriprionTxt += $"_{val}";
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
            void FillColor(List<short> list, RibbonCombo combo)
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
        private void PolylineTo2Polylines(ObjectId id, double offset, ref Autodesk.AutoCAD.DatabaseServices.Polyline polyline_1, ref Autodesk.AutoCAD.DatabaseServices.Polyline polyline_2)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // получаем полилинию по Id
                    var polyline = (Autodesk.AutoCAD.DatabaseServices.Polyline)tr.GetObject(id, OpenMode.ForRead);

                    // проверка на то, что полилиния не является осью
                    var axis = TsoddHost.Current.axis.FirstOrDefault(a => a.PolyHandle == polyline.Handle);
                    if (axis != null)
                    {
                        MessageBox.Show($" Ошибка создания разметки. Данная полилиния уже была выбрана для оси \" {axis.Name} \" ");
                        tr.Abort();
                        return;
                    }

                    // считаем offset
                    offset = db.Insunits == UnitsValue.Millimeters ? offset * 1000 : offset;

                    // пробуем получить подобные полилинии
                    try { polyline_1 = (Autodesk.AutoCAD.DatabaseServices.Polyline)polyline.GetOffsetCurves(offset / 2)[0]; }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex) { ed.WriteMessage($"Не получилось создать подобную линию для полилинии с Id = {id}"); }

                    try { polyline_2 = (Autodesk.AutoCAD.DatabaseServices.Polyline)polyline.GetOffsetCurves(-offset / 2)[0]; }
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
                    catch { outId = ObjectId.Null; return; }  // если объекта нет, то выходим

                    TsoddXdataElement tsoddXdataElement = new TsoddXdataElement();
                    tsoddXdataElement.Parse(id);

                    // если  есть и master и slave полилиния, то понимаем, что объект был полилиние, а значит пора преобразовать в одну полилинию
                    if (tsoddXdataElement.MasterPolylineID != ObjectId.Null)
                    {
                        // удаляем старый текст
                        if (tsoddXdataElement.MtextID != ObjectId.Null)
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
                            var masterPolyline = (Autodesk.AutoCAD.DatabaseServices.Polyline)tr.GetObject(tsoddXdataElement.MasterPolylineID, OpenMode.ForRead);        // master полилиния
                            var slavePolyline = (Autodesk.AutoCAD.DatabaseServices.Polyline)tr.GetObject(tsoddXdataElement.SlavePolylineID, OpenMode.ForRead);        // master полилиния

                            // расстояние между линиями 
                            double dist = masterPolyline.StartPoint.DistanceTo(slavePolyline.StartPoint);

                            // оффнтим masterPolyline на половину полученного расстояния и проверяем сторону оффсета
                            Autodesk.AutoCAD.DatabaseServices.Polyline tempPoliline_1 = null;
                            Autodesk.AutoCAD.DatabaseServices.Polyline tempPoliline_2 = null;

                            try { tempPoliline_1 = (Autodesk.AutoCAD.DatabaseServices.Polyline)masterPolyline.GetOffsetCurves(dist / 2)[0]; }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex) { ed.WriteMessage($"Не получилось создать подобную линию для полилинии с Id = {id}"); }

                            try { tempPoliline_2 = (Autodesk.AutoCAD.DatabaseServices.Polyline)masterPolyline.GetOffsetCurves(-dist / 2)[0]; }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex) { ed.WriteMessage($"Не получилось создать подобную линию для полилинии с Id = {id}"); }

                            // проверка расстояния
                            double dist_1 = tempPoliline_1.StartPoint.DistanceTo(slavePolyline.StartPoint);
                            double dist_2 = tempPoliline_2.StartPoint.DistanceTo(slavePolyline.StartPoint);

                            Autodesk.AutoCAD.DatabaseServices.Polyline resultPolyline = dist_1 < dist_2 ? tempPoliline_1 : tempPoliline_2;

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


        /// <summary>
        ///  Метод  инвертирующий линии у мультилинии
        /// </summary>
        public void LineTypeInvert()
        {
            var listElementsID = RibbonInitializer.Instance.GetAutoCadSelectionObjectsId(new List<string> { "LWPOLYLINE" });

            if (listElementsID == null) return;

            HashSet<ObjectId> listOfReverseObjects = new HashSet<ObjectId>();

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    //проходимся по всем элементам выбора
                    foreach (var id in listElementsID)
                    {
                        TsoddXdataElement txde = new TsoddXdataElement();
                        txde.Parse(id);

                        if (txde.Type != TsoddElement.Mline)
                        {
                            ed.WriteMessage("\n Объект разметки не подходит для команды инвертирования. \n");
                            continue;
                        }

                        var masterPoly = (Autodesk.AutoCAD.DatabaseServices.Polyline)tr.GetObject(txde.MasterPolylineID, OpenMode.ForWrite);
                        var slavePoly = (Autodesk.AutoCAD.DatabaseServices.Polyline)tr.GetObject(txde.SlavePolylineID, OpenMode.ForWrite);

                        if (masterPoly != null && slavePoly != null)
                        {

                            // проверка на то, что объект уже был изменен
                            if (listOfReverseObjects.TryGetValue(txde.MasterPolylineID, out ObjectId val)) continue;

                            // сохраняем во временную переменную данные
                            (string lineType, double width, int color) tempProp = (string.Empty, 0, 0);
                            tempProp.lineType = masterPoly.Linetype;
                            tempProp.width = masterPoly.ConstantWidth;
                            tempProp.color = masterPoly.ColorIndex;

                            masterPoly.Linetype = slavePoly.Linetype;
                            masterPoly.ConstantWidth = slavePoly.ConstantWidth;
                            masterPoly.ColorIndex = slavePoly.ColorIndex;

                            slavePoly.Linetype = tempProp.lineType;
                            slavePoly.ConstantWidth = tempProp.width;
                            slavePoly.ColorIndex = tempProp.color;

                            listOfReverseObjects.Add(txde.MasterPolylineID);
                        }
                    }

                    tr.Commit();
                }
            }
        }



        /// <summary>
        ///  Метод  инвертирующий положение текста у LineType
        /// </summary>
        public void LineTypeTextInvert()
        {
            var listElementsID = RibbonInitializer.Instance.GetAutoCadSelectionObjectsId(new List<string> { "LWPOLYLINE" });

            if (listElementsID == null) return;

            HashSet<ObjectId> listOfReverseObjects = new HashSet<ObjectId>();

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.Polyline masterPoly = null;
                    Autodesk.AutoCAD.DatabaseServices.Polyline slavePoly = null;
                    Autodesk.AutoCAD.DatabaseServices.Polyline basePoly = null;
                    Autodesk.AutoCAD.DatabaseServices.Polyline resultPoly = null;
                    Autodesk.AutoCAD.DatabaseServices.Polyline tempPoly_1 = null;
                    Autodesk.AutoCAD.DatabaseServices.Polyline tempPoly_2 = null;

                    double offset = 0;
                    double dist_1 = double.MaxValue;
                    double dist_2 = double.MaxValue;

                    double scale = db.Cannoscale.DrawingUnits;
                   //double scaleFactor = db.Insunits == UnitsValue.Millimeters ? 0.001 : 1;

                    //проходимся по всем элементам выбора
                    foreach (var id in listElementsID)
                    {
                        TsoddXdataElement txde = new TsoddXdataElement();
                        txde.Parse(id);

                        if (txde.Type == TsoddElement.Axis) continue;   // если это ось - то пропускаем

                        if (txde.Type == TsoddElement.Line || txde.Type == TsoddElement.Mline)             // если это простая
                        {
                            // проверка на то, что объект уже был изменен
                            if (listOfReverseObjects.TryGetValue(txde.MasterPolylineID, out ObjectId val)) continue;

                            MText mtxt;
                            try { mtxt = (MText)tr.GetObject(txde.MtextID, OpenMode.ForWrite); }
                            catch { ed.WriteMessage("\n Ошибка объекта mtext (данные полилинии разметки) \n"); return; }

                            offset = 0;
                            masterPoly = (Autodesk.AutoCAD.DatabaseServices.Polyline)tr.GetObject(txde.MasterPolylineID, OpenMode.ForRead);
                            basePoly = masterPoly;

                            if (txde.SlavePolylineID != ObjectId.Null)
                            {
                                slavePoly = (Autodesk.AutoCAD.DatabaseServices.Polyline)tr.GetObject(txde.SlavePolylineID, OpenMode.ForRead);
                                offset = masterPoly.StartPoint.DistanceTo(slavePoly.StartPoint) / 2;

                                try { tempPoly_1 = (Autodesk.AutoCAD.DatabaseServices.Polyline)masterPoly.GetOffsetCurves(offset)[0]; }
                                catch (Autodesk.AutoCAD.Runtime.Exception ex) { ed.WriteMessage($"Не получилось создать подобную линию для полилинии с Id = {id}"); }

                                try { tempPoly_2 = (Autodesk.AutoCAD.DatabaseServices.Polyline)masterPoly.GetOffsetCurves(-offset)[0]; }
                                catch (Autodesk.AutoCAD.Runtime.Exception ex) { ed.WriteMessage($"Не получилось создать подобную линию для полилинии с Id = {id}"); }

                                // проверка расстояния
                                dist_1 = tempPoly_1.StartPoint.DistanceTo(slavePoly.StartPoint);
                                dist_2 = tempPoly_2.StartPoint.DistanceTo(slavePoly.StartPoint);

                                basePoly = dist_1 < dist_2 ? tempPoly_1 : tempPoly_2;
                            }

                            // пользовательские настр-йки
                            var userOptions = JsonReader.LoadFromJson<Options>(FilesLocation.JsonOptionsPath);

                            try { tempPoly_1 = (Autodesk.AutoCAD.DatabaseServices.Polyline)basePoly.GetOffsetCurves(userOptions.LineTypeTextHeight * scale + offset)[0]; }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex) { ed.WriteMessage($"Не получилось создать подобную линию для полилинии с Id = {id}"); }

                            try { tempPoly_2 = (Autodesk.AutoCAD.DatabaseServices.Polyline)basePoly.GetOffsetCurves(-userOptions.LineTypeTextHeight * scale - offset)[0]; }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex) { ed.WriteMessage($"Не получилось создать подобную линию для полилинии с Id = {id}"); }

                            // определяем результирующую линию
                            dist_1 = tempPoly_1.GetPointAtDist(tempPoly_1.Length / 2).DistanceTo(mtxt.Location);
                            dist_2 = tempPoly_2.GetPointAtDist(tempPoly_2.Length / 2).DistanceTo(mtxt.Location);

                            resultPoly = dist_1 > dist_2 ? tempPoly_1 : tempPoly_2;

                            Point3d centerPoint = resultPoly.GetPointAtDist(resultPoly.Length / 2);
                            double parametr = resultPoly.GetParameterAtPoint(centerPoint);
                            Vector3d tan = resultPoly.GetFirstDerivative(parametr).GetNormal();
                            double angle = Math.Atan2(tan.Y, tan.X);

                            if (dist_1 < dist_2) angle += Math.PI;

                            mtxt.Location = centerPoint;
                            mtxt.Rotation = angle;

                            listOfReverseObjects.Add(txde.MasterPolylineID);

                        }
                    }
                    tr.Commit();
                }
            }
        }


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

            // считаем offset
            double koef = db.Insunits == UnitsValue.Millimeters ? 1000 : 1;

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {

                    var polyline = (Autodesk.AutoCAD.DatabaseServices.Polyline)tr.GetObject(id, OpenMode.ForWrite);

                    // проверка на то, что полилиния не является осью
                    var axis = TsoddHost.Current.axis.FirstOrDefault(a => a.PolyHandle == polyline.Handle);
                    if (axis != null)
                    {
                        MessageBox.Show($" Ошибка создания разметки. Данная полилиния уже была выбрана для оси \" {axis.Name} \" ");
                        tr.Abort();
                        return false;
                    }

                    // Текущий масштаб CANNOSCALE
                    var ocm = db.ObjectContextManager;
                    var occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");
                    var curSC = occ.CurrentContext as AnnotationScale;

                  
                    polyline.Linetype = lineType.Name;





                    polyline.LinetypeScale = koef / curSC.DrawingUnits;
                    polyline.ConstantWidth = lineType.Width * koef;
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
            // пользовательские настр-йки
            var userOptions = JsonReader.LoadFromJson<Options>(FilesLocation.JsonOptionsPath);
   
            if (polylineID == null) return null;

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    // исходная полилиния
                    var polyline = (Autodesk.AutoCAD.DatabaseServices.Polyline)tr.GetObject(polylineID, OpenMode.ForRead);

                    // уточняем положение текста
                    Autodesk.AutoCAD.DatabaseServices.Polyline tempPoly = null;
                    double scale = db.Cannoscale.DrawingUnits;
                    double scaleFactor = db.Insunits == UnitsValue.Millimeters ? 0.001 : 1;

                    try { tempPoly = (Autodesk.AutoCAD.DatabaseServices.Polyline)polyline.GetOffsetCurves((userOptions.LineTypeTextHeight + 0.1) * scale)[0]; }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex) { }

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
                    mt.Contents = $"{lineTypeName} (" + $@"%<\AcObjProp Object(%<\_ObjId {mTextId}>%).Length \f ""%lu2%ct8[{scaleFactor}]"">%" + " м)";
                    mt.TextHeight = userOptions.LineTypeTextHeight * scale;
                    mt.Height = userOptions.LineTypeTextHeight;
                    mt.Attachment = AttachmentPoint.MiddleCenter;

                    // Получаем таблицу текстовых стилей
                    TextStyleTable textStyleTable = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);

                    // Перебираем все стили
                    foreach (ObjectId styleId in textStyleTable)
                    {
                        TextStyleTableRecord styleRecord = (TextStyleTableRecord)tr.GetObject(styleId, OpenMode.ForRead);
                        if (!string.IsNullOrEmpty(styleRecord.Name) && styleRecord.Name == userOptions.LineTypeTextStyle) mt.TextStyleId = styleId;
                    }

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


        // метод создает список RibbonButton для каждого кастомного типа линий
        public void ListOfMarksLinesLoad(int bmpWidth, int bmpHeight, string preSelect = null)
        {
            splitMarksLineTypes.Items.Clear();

            // проходимся по всем типам линий в файле
            foreach (var lineType in LineTypeReader.Parse())
            {
                if (lineType.IsMlineElement) continue;  //  если это составная часть мультилинии, то пропускаем 

                var bmp = new Bitmap(bmpWidth, bmpHeight); // новый Bitmap для отрисовки типа линии

                using (var g = Graphics.FromImage(bmp))     // холст созданный, по Bitmap
                {

                    using (var pen = new Pen(Brushes.WhiteSmoke, 2f))  //  Кисть для рисовния 
                    {
                        // прозрачный фон bmp 
                        g.Clear(Color.Transparent);

                        // подпись слоя
                        using (var font = new System.Drawing.Font("Mipgost", 10f, System.Drawing.FontStyle.Bold))
                        {
                            g.DrawString(lineType.Name, font, Brushes.WhiteSmoke, new PointF(13 - lineType.Name.Length * 2, 2));
                        }

                        // получаем паттерн штриховой линии простого типа лини
                        if (lineType is AcadLineType lt && lt.PatternValues.Count > 0)
                        {
                            Autodesk.AutoCAD.Colors.Color cl = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, lt.ColorIndex[0]);
                            pen.Color = cl.ColorValue;

                            float[] patternArray = GetPatternArray(lt.PatternValues[0], 5);
                            pen.DashPattern = patternArray;
                            int y = bmpHeight / 2 + 1;
                            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                            g.DrawLine(pen, 50, y, bmpWidth - 6, y);
                        }

                        // если это мультилиния, то находим все линии и берем паттерн штриховки у них
                        if (lineType is AcadMLineType mlt)
                        {
                            int y = bmpHeight;
                            foreach (var line in mlt.MLineLineTypes)
                            {
                                Autodesk.AutoCAD.Colors.Color cl = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, line.ColorIndex[0]);
                                pen.Color = cl.ColorValue;

                                float[] patternArray = GetPatternArray(line.PatternValues[0], 5);
                                pen.DashPattern = patternArray;
                                y -= bmpHeight / 3 + 1;
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
        public float[] GetPatternArray(List<double> lineTypeValues, float scale)
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
        public void FillBlocksMenu(RibbonSplitButton split, string tag, string preselect = null)
        {
            split.Items.Clear();

            //если это знаки, то будем получать список 
            string group = null;
            if (tag == "SIGN" && TsoddHost.Current.currentSignGroup != null) group = TsoddHost.Current.currentSignGroup;

            var listOfBlocks = TsoddBlock.GetListOfBlocks(tag, group);    // формируем список блоков по тегу

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
                            TsoddBlock.InsertStandOrMarkBlock(name, true);
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

        public List<ObjectId> GetAutoCadSelectionObjectsId(SelectionFilter filter, bool singleSelection = false)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            string textMessage = "\n Выберите объекты, можно рамкой выбрать несколько:";
            if (singleSelection) textMessage = "\n Выберите объект:";

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

        public List<ObjectId> GetAutoCadSelectionObjectsId(List<string> filterList, bool singleSelection = false)
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

            return GetAutoCadSelectionObjectsId(filter, singleSelection);
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
        /// Метод создает мультивыноску для объектов TSODD
        /// </summary>
        public void CreateMLeaderForTsoddObject()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // пользовательские настройки
            var userOptions = JsonReader.LoadFromJson<Options>(FilesLocation.JsonOptionsPath);
            double scale = db.Cannoscale.DrawingUnits;

            string txt = "";

            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\n Выберите объект (блок стойки, блок знака, блок разметки или линию разметки, Esc - выход) ";
            pso.SingleOnly = true;

            var filter = new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "INSERT,LWPOLYLINE,POLYLINE") });

            var psr = ed.GetSelection(pso, filter);

            if (psr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n Ошибка выбора объекта \n");
                return;
            }

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // узнаем тип выбранного объекта
                ObjectId id = psr.Value.GetObjectIds().First();

                DBObject dbo = (DBObject)tr.GetObject(id, OpenMode.ForRead);

                switch (dbo)
                {
                    case BlockReference blockRef:       // блок

                        foreach (ObjectId attId in blockRef.AttributeCollection)
                        {
                            var attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                            if (attRef.Tag.Equals("НОМЕР_ЗНАКА", StringComparison.OrdinalIgnoreCase)) { txt = attRef.TextString; break; }
                            if (attRef.Tag.Equals("NUMBER", StringComparison.OrdinalIgnoreCase)) { txt = attRef.TextString; break; }
                        }

                        // если не нашли подходящий тег, то просто возьмем имя блока
                        if (txt == "") txt = blockRef.Name;

                        break;

                    case Autodesk.AutoCAD.DatabaseServices.Polyline poly:

                        // пробуем прочитать xData
                        TsoddXdataElement txde = new TsoddXdataElement();
                        txde.Parse(poly.Id);

                        if (txde.Type == TsoddElement.Axis)     // если это ось
                        {
                            txt = TsoddHost.Current.axis.FirstOrDefault(a => a.PolyID == poly.Id).Name;
                            break;
                        }

                        // читаем xData master полилинии
                        if (txde.MasterPolylineID != ObjectId.Null)
                        {
                            double koef = db.Insunits == UnitsValue.Meters ? 1 : 0.001;
                            txde.Parse(txde.MasterPolylineID);
                            string mTextId = txde.MasterPolylineID.OldId.ToString();
                            txt = $"{txde.Number} (" + $@"%<\AcObjProp Object(%<\_ObjId {mTextId}>%).Length \f ""%lu2%ct8[{koef}]"">%" + " м)";
                        }
                        // если не нашли подходящую xData
                        if (txt == "") txt = "не определили имя";

                        break;
                }

                Point3d arrowPoint = Point3d.Origin;
                Point3d textPoint = Point3d.Origin;

                // точка стрелки
                var peo1 = new PromptPointOptions("\n Выберете позицию стрелки мультивыноски: ");
                var per1 = ed.GetPoint(peo1);
                if (per1.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\n Ошибка построения мультивыноски  \n");
                    return;
                }

                arrowPoint = per1.Value;

                // точка стрелки
                var peo2 = new PromptPointOptions("\n Выберете позицию текста мультивыноски: ");
                var per2 = ed.GetPoint(peo2);
                if (per2.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\n Ошибка построения мультивыноски \n");
                    return;
                }

                textPoint = per2.Value;

                MText mtext = new MText
                {
                    Contents = txt,
                    TextHeight = userOptions.MleaderTextHeight * scale,
                    Location = textPoint,
                    Attachment = AttachmentPoint.MiddleCenter,
                    Annotative = AnnotativeStates.True
                };

                MLeader mleader = new MLeader
                {
                    MText = mtext,
                    ContentType = ContentType.MTextContent,
                    ArrowSize = 2,
                    LandingGap = 1.0
                };

                int leaderIndex = mleader.AddLeader();
                int lineIndex = mleader.AddLeaderLine(leaderIndex);
                mleader.AddFirstVertex(lineIndex, arrowPoint);
                mleader.AddLastVertex(lineIndex, mtext.Location);


                DBDictionary dict = tr.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead) as DBDictionary;

                if (dict != null)
                {
                    foreach (DBDictionaryEntry entry in dict)
                    {
                        MLeaderStyle style = tr.GetObject(entry.Value, OpenMode.ForRead) as MLeaderStyle;
                        if (entry.Value != ObjectId.Null && style.Name == userOptions.MleaderStyle) mleader.MLeaderStyle = entry.Value;
                    }
                }

                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                btr.AppendEntity(mleader);
                tr.AddNewlyCreatedDBObject(mleader, true);
                tr.Commit();
                ed.Regen();
            }
        }

        /// <summary>
        ///  Получает ПК оси указанной точки
        /// </summary>
        public void GetPkOnAxis()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
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
                ed.PointMonitor -= OnPointMonitor;
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
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
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
