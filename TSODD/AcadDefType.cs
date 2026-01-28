using ACAD_test;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
//using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;


namespace TSODD
{
    public enum AcadDefType { LineType, MlineType }
    public interface IAcadDef
    {
        string Name { get; set; }           // наименование
        string Description { get; set; }    // описание
        AcadDefType Type { get; }           // тип объекта
        void ParseValue(string line);      // метод парсинга значений
        bool IsReadyForAdding();            // метод, проверяет заполнены ли данные по типу линии или мультилинии
        bool IsMlineElement { get; set; }

    }

    // тип линии
    public class AcadLineType : IAcadDef
    {
        public string Name { get; set; } = "";
        public string Description { get; set; } = "";
        public AcadDefType Type => AcadDefType.LineType;
        public bool IsMlineElement { get; set; } = false;
        public List<List<double>> PatternValues { get; set; } = new List<List<double>>();
        public List<short> ColorIndex { get; set; } = new List<short>();
        public List<double> Width { get; set; } = new List<double>();

        // метод парсинга значений
        public void ParseValue(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return;
            char startChar = ' ';
            string[] split = null;
            string stringValue = null;

            startChar = line[0];    // первый символ в строке

            switch (startChar)
            {
                case '*':   // Строка наименования и описания обычного типа линии

                    line = line.Substring(1).TrimStart();      // убираем "*"
                    split = line.Split(',');

                    if (split.Count() == 2)     // верный формат 
                    {
                        Name = split[0].Trim();                 // имя типа линии
                        Description = split[1].Trim();          // описание типа линии
                        if (Description.Contains("multiLine")) IsMlineElement = true;
                    }

                    break;

                case 'A':   // Cтрока параметров обычного типа линии

                    line = line.Substring(1).TrimStart();                                // убираем символ 'A'
                    if (line.StartsWith(",")) line = line.Substring(1).TrimStart();      // убираем запятую

                    split = line.Split(',');

                    List<double> pattern = new List<double>();

                    foreach (var s in split)
                    {
                        stringValue = s.Trim();
                        if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var doubleValue))
                        {
                            pattern.Add(doubleValue);
                        }
                    }

                    PatternValues.Add(pattern);

                    break;

                case 'C':   // Строка параметров цвета

                    line = line.Substring(1).TrimStart();                                // убираем символ 'С'
                    if (line.StartsWith(",")) line = line.Substring(1).TrimStart();      // убираем запятую

                    split = line.Split(',');

                    foreach (var s in split)
                    {
                        stringValue = s.Trim();
                        if (short.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var intValue))
                        {
                            ColorIndex.Add(intValue);
                        }
                    }
                    break;

                case 'W':   // Строка параметров 

                    line = line.Substring(1).TrimStart();                                // убираем символ 'W'
                    if (line.StartsWith(",")) line = line.Substring(1).TrimStart();      // убираем запятую

                    split = line.Split(',');

                    foreach (var s in split)
                    {
                        stringValue = s.Trim();
                        if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var doubleValue))
                        {
                            Width.Add(doubleValue);
                        }
                    }
                    break;
            }
        }

        public void CopyLineType(IAcadDef acadLineType)
        {
            if (acadLineType is AcadLineType lt)
            {
                Description = lt.Description;
                PatternValues = lt.PatternValues;
                ColorIndex = lt.ColorIndex;
                IsMlineElement = lt.IsMlineElement;
                Width = lt.Width;
            }
        }

        public bool IsReadyForAdding()
        {
            return !string.IsNullOrEmpty(Name) &&
                   !string.IsNullOrEmpty(Description) &&
                   PatternValues.Count > 0 &&
                   ColorIndex.Count > 0 &&
                   Width.Count > 0;
        }
    }


    // тип мультилинии
    public class AcadMLineType : IAcadDef
    {
        public string Name { get; set; } = "";
        public string Description { get; set; } = "";
        public AcadDefType Type => AcadDefType.MlineType;
        public List<double> Offset { get; set; } = new List<double>();
        public List<AcadLineType> MLineLineTypes { get; set; } = new List<AcadLineType>();
        public bool IsMlineElement { get; set; } = false;

        // метод парсинга значений
        public void ParseValue(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return;
            char startChar = ' ';
            string[] split = null;
            string stringValue = null;

            startChar = line[0];    // первый символ в строке

            switch (startChar)
            {
                case '#':   // Строка наименования и описания обычного типа линии

                    line = line.Substring(1).TrimStart();      // убираем "*"
                    split = line.Split(',');

                    if (split.Count() == 2)     // верный формат 
                    {
                        Name = split[0].Trim();             // имя типа линии
                        Description = split[1].Trim();      // описание типа линии
                    }

                    break;

                case 'L':   // Cтрока параметров обычного типа линии

                    line = line.Substring(1).TrimStart();                                // убираем символ 'L'
                    if (line.StartsWith(",")) line = line.Substring(1).TrimStart();      // убираем запятую

                    MLineLineTypes.Add(new AcadLineType { Name = line.Trim() });     // временно сосздаем тип линии с соответствующим именем

                    break;

                case 'O':

                    line = line.Substring(2).TrimStart();                                // убираем символ 'O'
                    if (line.StartsWith(",")) line = line.Substring(1).TrimStart();      // убираем запятую

                    split = line.Split(',');

                    foreach (var s in split)
                    {
                        stringValue = s.Trim();
                        if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var doubleValue))
                        {
                            Offset.Add(doubleValue);
                        }
                    }
                    break;
            }
        }

        public bool IsReadyForAdding()
        {
            bool val_1 = !string.IsNullOrEmpty(Name) &&
                   !string.IsNullOrEmpty(Description) &&
                   MLineLineTypes.Count > 0 &&
                   Offset.Count > 0;

            bool val_2 = true;
            foreach (var lineType in MLineLineTypes)
            {
                if (lineType.IsReadyForAdding() == false) val_2 = false;
            }

            return val_1 && val_2;
        }
    }

    public class CurrentLineType
    {
        public string Name { get; set; } = "";
        public double Width { get; set; } = 0;
        public short ColorIndex { get; set; } = 0;
    }

    public static class LineTypeReader
    {

        // метод парсит список типов линий
        public static List<IAcadDef> Parse()
        {
            List<IAcadDef> result = new List<IAcadDef>();

            if (!File.Exists(FilesLocation.linPath))
            {
                MessageBox.Show("Ошибка чтения файла с типами линий");
                return null;
            }

            IAcadDef current = null;

            foreach (var raw in File.ReadLines(FilesLocation.linPath))
            {
                var line = raw.Trim();
                if (string.IsNullOrEmpty(line)) continue;       // пропускаем пустыке строки

                // Строка наименования и описания обычного типа линии
                if (line.StartsWith("*", StringComparison.Ordinal)) current = new AcadLineType();

                // Строка наименования и описания типа мультилинии
                if (line.StartsWith("#", StringComparison.Ordinal)) current = new AcadMLineType();

                current.ParseValue(line);

                if (line.StartsWith("W", StringComparison.Ordinal) || line.StartsWith("O", StringComparison.Ordinal))
                {
                    // последняя строка описания типа линии или мультилинии - пора добавить в список
                    result.Add(current);
                }
            }

            // в мультилиниях надо подменить имена на типы линий
            foreach (var multiLine in result)
            {
                if (multiLine.Type == AcadDefType.LineType) continue; // пропускаем простые линии

                if (multiLine is AcadMLineType ml)
                {
                    foreach (var line in ml.MLineLineTypes)
                    {
                        IAcadDef acadLineType = result.FirstOrDefault(l => l.Name == line.Name);
                        if (acadLineType != null)
                        {
                            line.CopyLineType(acadLineType);
                        }
                    }
                }
            }
            return result;
        }


        // метод добавляет тип линии или мультилинии в БД
        public static void AddLineTypeToBD(IAcadDef acadLineType, bool messageOK = false)
        {

            // если параметры линии не заполнены, то выходим
            if (acadLineType.IsReadyForAdding() == false)
            {
                MessageBox.Show("Не получилось добавить тип линии в БД. Отсутствуют данные типа линии");
                return;
            }

            // получим текущий список типов линий
            var listOfLineTypes = Parse();

            // проверка дубликатов
            var dublicate = listOfLineTypes.Any(l => l.Name == acadLineType.Name);
            if (dublicate)
            {
                MessageBox.Show($"Тип линии с наименованием \"{acadLineType.Name}\" уже есть в БД");
                return;
            }

            // формируем строку записи в файл .lin
            string txt = "";

            if (acadLineType is AcadLineType lt)        // если это простой тип линии
            {
                // имя и описание
                txt += "\n" + $"*{lt.Name},{lt.Description}";

                foreach (var pattern in lt.PatternValues)
                {
                    // параметры линии
                    txt += "\nA";
                    foreach (var val in pattern) txt += $",{val}";
                }

                // Цвета линии
                txt += $"\nC";
                foreach (var color in lt.ColorIndex) txt += $",{color}";

                // Толщины линии 
                txt += $"\nW";
                foreach (var width in lt.Width) txt += $",{width}";

                txt += "\n";    // пропуск строки
            }

            if (acadLineType is AcadMLineType mlt)        // если это простой тип линии
            {
                // имя и описание
                txt += "\n" + $"#{mlt.Name},{mlt.Description}";

                foreach (var lineType in mlt.MLineLineTypes)
                {
                    // имя типа линии
                    txt += $"\nL,{lineType.Name}";
                }

                // Отступы (расстояния между линиями)
                txt += $"\nO";
                foreach (var offset in mlt.Offset) txt += $",{offset}";

                txt += "\n";    // пропуск строки
            }

            // запись в файл
            File.AppendAllText(FilesLocation.linPath, txt, new UTF8Encoding(false));

            // обновляем типы линии в автокаде
            RefreshLineTypesInAcad();

            // обновляем типы линии на ribbon
            RibbonInitializer.Instance.ListOfMarksLinesLoad(200, 20);

            if (messageOK) MessageBox.Show($"Тип линии с наименованием \"{acadLineType.Name}\" успешно добавлен в БД");
        }


        // метод удаляет объект типа линии или мультилинии
        public static void DeleteLineTypeFromBDs(string name)
        {
            // поиск имени типа линии в файле
            // получим текущий список типов линий
            var listOfLineTypes = Parse();

            // проверка дубликатов
            IAcadDef searchVal = null;
            IAcadDef lineType = listOfLineTypes.FirstOrDefault(l => l.Name == name);
            if (lineType == null)
            {
                MessageBox.Show($"Не найден тип линии с именем \"{name}\".");
                return;
            }

            // если это простой тип линии
            if (lineType is AcadLineType al)
            {
                // проверка на то, что тип линии используется 
                foreach (var acadDef in listOfLineTypes)
                {
                    if (acadDef is AcadMLineType acadMline)
                    {
                        if (acadMline.MLineLineTypes.Any(l => l.Name == name))
                        {
                            MessageBox.Show($"Ошибка удаления типа линии. Тип линии \"{name}\" используется в двойной линии \"{acadMline.Name}\".");
                            return;
                        }
                    }
                }

                if (DeleteLineTypeFromAcad(name) == false)      // беда - не получилось удалить 
                {
                    MessageBox.Show($"Не получилось удалить тип линии \"{name}\". Вероятно он используется в таких элементах как блок или внешняя ссылка.");
                    return;
                }

                searchVal = listOfLineTypes.FirstOrDefault(l => l.Name == name);
                if (searchVal != null) listOfLineTypes.Remove(searchVal);
            }

            // если это тип мультилинии
            if (lineType is AcadMLineType aml)
            {
                searchVal = listOfLineTypes.FirstOrDefault(l => l.Name == name);
                if (searchVal != null) listOfLineTypes.Remove(searchVal);
            }

            // очищаем текущий файл типов линий
            File.WriteAllText(FilesLocation.linPath, string.Empty);
            foreach (var obj in listOfLineTypes)
            {
                AddLineTypeToBD(obj);
            }
        }


        // метод обновляет типы линии в автокаде
        public static void RefreshLineTypesInAcad()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            // список всех типов линий (кастомных конечно)
            var listOfLineTypes = Parse();

            // список со всеми типами линий, учитывая паттерны штриховки
            List<string> tempList = new List<string>();
            List<string> tempNameList = new List<string>();
            StringBuilder sb = new StringBuilder();
            bool dublicate = true;

            foreach (var lineType in listOfLineTypes)
            {
                sb.Clear();

                if (lineType is AcadLineType lt)
                {
                    foreach (var listVal in lt.PatternValues)
                    {
                        string txt_val = "";
                        foreach (var val in listVal) txt_val += $",{val}";
                        string name = $"{lt.Name} ({txt_val.Replace(",", "_").Substring(1).TrimStart()})";

                        // проверка на сплошную линию
                        if (listVal.Contains(1000)) name = $"{lt.Name} (continuous)";

                        dublicate = tempNameList.Any(n => n == name);

                        if (!dublicate)
                        {
                            tempNameList.Add(name);
                            sb.AppendLine($"*{name},{lt.Description}");
                            sb.AppendLine($"A{txt_val} \n");
                        }
                    }

                    tempList.Add(sb.ToString());
                }
            }

            // запись всех типов линий с учетом всех паттернов штриховки
            File.WriteAllLines(FilesLocation.separatedLinPath, tempList);

            foreach (var lineType in tempNameList)
            {
                using (doc.LockDocument())
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var ltt = (LinetypeTable)tr.GetObject(db.LinetypeTableId, OpenMode.ForRead);

                        if (!ltt.Has(lineType)) db.LoadLineTypeFile(lineType, FilesLocation.separatedLinPath);    // грузим тип линии в автокад
                        tr.Commit();
                    }
                }
            }
        }


        public static bool DeleteLineTypeFromAcad(string name)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            using (doc.LockDocument())

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ltt = (LinetypeTable)tr.GetObject(db.LinetypeTableId, OpenMode.ForRead);

                using (doc.LockDocument())
                {
                    if (!ltt.Has(name))
                    {
                        tr.Commit();
                        return true; // если такой тип линии не был подгружен, то и нечего удалять
                    }

                    //  если текущий тип равен удаляемому — переключим на CONTINUOUS
                    var cur = (string)Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CELTYPE");
                    if (string.Equals(cur, name, System.StringComparison.OrdinalIgnoreCase))
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CELTYPE", "CONTINUOUS");

                    // пробуем переназначить у объектов тип линии
                    var ltrIdToDelete = ltt[name];
                    var continuousId = ltt["CONTINUOUS"];

                    var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    foreach (ObjectId layerId in layerTable)
                    {
                        var temp_ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                        if (temp_ltr.LinetypeObjectId == ltrIdToDelete)
                        {
                            temp_ltr.UpgradeOpen();
                            temp_ltr.LinetypeObjectId = continuousId;
                        }
                    }

                    //  Объекты с LineType на уровне объекта 
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    foreach (ObjectId btrId in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                        foreach (ObjectId entId in btr)
                        {
                            if (!entId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Entity))))
                                continue;

                            var ent = (Entity)tr.GetObject(entId, OpenMode.ForRead);
                            if (ent.LinetypeId == ltrIdToDelete)
                            {
                                ent.UpgradeOpen();
                                ent.LinetypeId = continuousId;
                            }
                        }
                    }

                    // пургеним тип линии
                    var ids = new ObjectIdCollection(new[] { ltrIdToDelete });
                    db.Purge(ids);
                    if (ids.Count == 0)            // всё ещё где-то используется (на объектах и т.п.) беда, придется пользователю вручную удалять
                    {
                        tr.Commit();
                        return false;
                    }


                    // ну и удаляем из таблицы автокада типов линий
                    var ltr = (LinetypeTableRecord)tr.GetObject(ltrIdToDelete, OpenMode.ForWrite);
                    ltt.UpgradeOpen();
                    ltr.Erase(true);
                    tr.Commit();
                    return true;
                }
            }
        }


        public static ObjectId[] CollectMarkLineTypeID(bool onlyMaster)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            SelectionFilter filter;
            if (onlyMaster)
            {
                filter = new SelectionFilter(new TypedValue[]
                      {
                      new TypedValue((int)DxfCode.ExtendedDataRegAppName,AutocadXData.AppName),
                      new TypedValue((int)DxfCode.ExtendedDataAsciiString,"master")
                      });
            }
            else
            {
                filter = new SelectionFilter(new TypedValue[]
                      {
                      new TypedValue((int)DxfCode.ExtendedDataRegAppName,AutocadXData.AppName),
                      new TypedValue((int)DxfCode.ExtendedDataAsciiString,"master,slave")
                      });
            }

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // поиск с учетом фильтра
                var selection = ed.SelectAll(filter);

                if (selection.Status != PromptStatus.OK)    // неудачный поиск
                {
                    ed.WriteMessage("\n В чертеже не найдено разметки в виде линий \n");
                    return new ObjectId[0];
                }
                else
                {
                    return selection.Value.GetObjectIds();
                }
            }

        }

    }
}
