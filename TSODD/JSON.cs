using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace TSODD
{
    internal static class JsonReader
    {
        // метод сохранения в json
        public static void SaveToJson<T>(T obj, string path)
        {
            string json = JsonSerializer.Serialize(obj, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }

        // метод загрузки из json
        public static T LoadFromJson<T>(string path)
        {
            string json = File.ReadAllText(path);
            var obj = JsonSerializer.Deserialize<T>(json);
            return obj;
        }

        // helper для сохранения имен в шапке таблиц и 
        private static void CreateTableHeader()
        {

            // шапка для таблицы со знаками
            List<string> signsHeader = new List<string>
            {
                "№_1_1_n",
                "Участок УДС_1_2_n",
                "Номер по ГОСТ Р 52290-2004_1_3_n",
                "Наименование_1_4_n",
                "Типоразмер_1_5_n",
                "Месторасположения_1_6_n",
                "Расположение_1_7_n",
                "Количество_1_8_n",
                "Наличие_1_9_n",
                "Площадь_1_10_n"
            };
            // серриализуем в json
            SaveToJson(signsHeader, FilesLocation.JsonTableHeaderSignsPath);


            // шапка для таблицы со знаками
            List<string> marksHeader = new List<string>
            {
                "№_1_1_v_2",
                "Участок УДС_1_2_v_2",
                "Номер по ГОСТ Р 51256-2018_1_3_v_2",
                "Протяженность/количество единиц_1_4_v_2",
                "Площадь_1_5_v_2",
                "Месторасположения_1_6_h_2",
                "Начало_2_6_n",
                "Конец_2_7_n",
                "Расположение_1_8_v_2",
                "Материал изготовления_1_9_v_2",
                "Наличие_1_10_v_2"
            };

            // серриализуем в json
            SaveToJson(marksHeader, FilesLocation.JsonTableHeaderMarksPath);


            // словарь для имен знаков
            Dictionary<string, string> signNames = new Dictionary<string, string>();

            // открываем excel файл и заполняем словарь
            using (var package = new ExcelPackage(new FileInfo(FilesLocation.ExcelTableSignsPath)))
            {
                var ws = package.Workbook.Worksheets[0];     // лист
                var last_row = ws.Dimension.End.Row;         // последняя строка

                for (int i = 1; i < last_row + 1; i++)
                {
                    var cur_num = ws.Cells[i, 3].Value == null ? "" : ws.Cells[i, 3].Value.ToString();
                    var cur_name = ws.Cells[i, 4].Value == null ? "" : ws.Cells[i, 4].Value.ToString();
                    if (cur_num != "" && !signNames.ContainsKey(cur_num)) signNames.Add(cur_num, cur_name);    // заполняем словарь
                }
            }
            SaveToJson<Dictionary<string, string>>(signNames, FilesLocation.JsonTableNamesSignsPath);
        }

    }
}
