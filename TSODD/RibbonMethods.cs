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
using Autodesk.AutoCAD.GraphicsInterface;


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
                db.ObjectErased += Db_ObjectErased;     // событие на удаление объекта из DB
                db.BeginSave += Db_BeginSave; ;
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
        private void Db_BeginSave(object sender, DatabaseIOEventArgs e)
        {
            var file = e.FileName;
            var extension = file.Split('.')[1];
            bool autoSave = extension.Contains("$");

            // если это не автосейв, то обновим данные 
            if (!autoSave)
            {
                SaveTsoddChanges();
            }
        }


        // обработчик удаления элемента из базы autoCAD
        private void Db_ObjectErased(object sender, Autodesk.AutoCAD.DatabaseServices.ObjectErasedEventArgs e)
        {
            if (!e.Erased) return;

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

                    // удаляем блок
                    TsoddBlock.DeleteAxisBlock(erasedObject);

                }
            }
        }

        // lifeguard для осей. после выполнения команды, проверяем есть ли в структуре _dontDeleteMe элементы, которые нужно спасти
        private void MdiActiveDocument_CommandEnded(object sender, CommandEventArgs e)
        {
            if (_dontDeleteMe.Count == 0) return;

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

                tr.Commit();
            }
            _dontDeleteMe.Clear();
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




        private void ApplyLineType(string name)
        {
            if (name == null) return;

            // проверяем изменения
            CheckSplitMarksLineTypesCurrentChanged(name);

            // парсим все типы линий
            var lineTypeList = LineTypeReader.Parse();

            // текущий тип линии
            var currentIAcadDef = lineTypeList.FirstOrDefault(lt => lt.Name == name);

            if (currentIAcadDef == null) return;

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
            }

            // заполняем первую линию
            // паттерн
            var patternLabel_1 = comboLineTypePattern_1.Current as RibbonLabel;
            var lineTypePattern_1 = patternLabel_1.Description;
            // имя типа линии в БД Autocad
            currentLineType_1.Name = $"{name} {lineTypePattern_1}";
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
                currentLineType_2.Name = $"{name} {lineTypePattern_2}";
                // толщина 
                var widthLabel_2 = comboLineTypeWidth_2.Current as RibbonLabel;
                Double.TryParse(widthLabel_2.Description, out double val2_1);
                if (val2_1 > 0) currentLineType_2.Width = val2_1;
                // цвет
                var colorLabel_2 = comboLineTypeColor_2.Current as RibbonLabel;
                short.TryParse(colorLabel_2.Description, out short val2_2);
                if (val2_2 > 0) currentLineType_2.ColorIndex = val2_2;
            }
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
                    comboLineTypePattern_2.IsEnabled = false;
                    comboLineTypeWidth_2.IsEnabled = false;
                    comboLineTypeColor_2.IsEnabled = false;
                    comboLineTypeOffset.IsEnabled = false;

                    // заполняем паттерны
                    FillPattern(lt.PatternValues, comboLineTypePattern_1);
                    // заполняем толщины
                    FillWidth(lt.Width, comboLineTypeWidth_1);
                    // заполняем цвета
                    FillColor(lt.ColorIndex, comboLineTypeColor_1);

                    break;

                case var l when l is AcadMLineType mlt:

                    // разблокируем контролы для второй линии
                    comboLineTypePattern_2.IsEnabled = true;
                    comboLineTypeWidth_2.IsEnabled = true;
                    comboLineTypeColor_2.IsEnabled = true;
                    comboLineTypeOffset.IsEnabled = true;

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
                            // заполняем расстояния между линиями
                            FillOffset(mlt.Offset, comboLineTypeOffset);
                        }
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



        // метод перестроения и сохранения объектов TSODD
        private void SaveTsoddChanges()
        {
            // обновим блоки осей
            foreach (Axis axis in TsoddHost.Current.axis)
            {
                TsoddBlock.AddAxisBlock(axis);
            }
        }



        // метод создает список RibbonButton для каждого кастомного типа линий
        private void ListOfMarksLinesLoad(int bmpWidth, int bmpHeight)
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
