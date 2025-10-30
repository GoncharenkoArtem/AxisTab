
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows.Documents;
using Autodesk.AutoCAD.Interop;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using TSODD.Properties;
using System.Windows.Forms;
using ACAD = Autodesk.AutoCAD.ApplicationServices;
using System.Runtime.Remoting.Messaging;
using System;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using Autodesk.AutoCAD.DatabaseServices;
using System.Data.Common;

namespace TSODD
{
    public class ExportExcel
    {
        EPPlusLicense _license;

        DirectoryInfo _settingsFolder;
        string _jsonTableNamesSignsPath;
        string _jsonTableHeaderSignsPath;
        string _jsonTableHeaderMarksPath;



        public ExportExcel()
        {
            _license = new EPPlusLicense();
            _license.SetNonCommercialOrganization("abc");

            var dllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            _settingsFolder = Directory.CreateDirectory(System.IO.Path.Combine(dllPath, "Support"));
            _jsonTableNamesSignsPath = Path.GetFullPath(Path.Combine(_settingsFolder.FullName, "NamesSigns.json"));
            _jsonTableHeaderSignsPath = Path.GetFullPath(Path.Combine(_settingsFolder.FullName, "TableHeaderSigns.json"));
            _jsonTableHeaderMarksPath = Path.GetFullPath(Path.Combine(_settingsFolder.FullName, "TableHeadersMarks.json"));


        }










        static void TableSignsNames()
        {



        }












        // метод эксопрта адресной ведомости знаков в эксель
        public void ExportSigns()
        {
            var exportFileName = $"Адресная ведомость дорожных знаков {ACAD.Application.DocumentManager.MdiActiveDocument.IsNamedDrawing}";

            SaveFileDialog saveFileDialog = new SaveFileDialog // создаем новое окно сохранения
            {
                Title = "Сохранить файл Excel",
                Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*",
                DefaultExt = "xlsx",
                FileName = exportFileName
            };
      
            // показываем окно
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;

                if (File.Exists(filePath))      // если файл существует, то проверим не открыт ли он
                {
                    FileStream fs = null;
                    try
                    {
                        fs = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                    }
                    catch (IOException)
                    {
                        MessageBox.Show("Ошибка сохранения ведомости. Файл открыт.");
                        return;
                    }
                    finally 
                    {
                       if(fs != null) fs.Dispose();
                    }
                }

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    // что бы там не было в файле - просто очищаем его в 0
                    while (package.Workbook.Worksheets.Count > 0)
                        package.Workbook.Worksheets.Delete(0);

                    // новый лист для ведомости
                    var ws = package.Workbook.Worksheets.Add("Дорожные знаки");

                    // выгружаем данные из JSON
                    var signHeaders = LoadFromJson<List<string>>(_jsonTableHeaderSignsPath);                // для шапки таблицы
                    var signNames = LoadFromJson<Dictionary<string, string>>(_jsonTableNamesSignsPath);     // имена знаков

                    // шапка таблицы
                    AddHeaders(ws, signHeaders);
                    int currentRow = 1;

                    List<Stand> stands = TsoddBlock.GetListOfRefBlocks<Stand>();             // список стоек файла
                    List<Sign> signs = TsoddBlock.GetListOfRefBlocks<Sign>(signNames);       // список знаков файла

                    // сортировка стоек по Осям
                    var standGroups = stands.GroupBy(g => g.AxisName);
                    foreach (var group in standGroups)
                    {
                        // сортировка по расстоянию
                        var orderedStands = group.OrderBy(s => s.Distance);

                        foreach (Stand stand in orderedStands)  // для каждой стойки найдем знаки
                        {
                            var splitSigns = signs.Where(s => s.StandHandle == stand.Handle);
                            foreach (Sign sing in splitSigns)
                            {
                                // пора записывать в ведомость
                                AddSignDataToTable(ws, stand, sing, ref currentRow);
                            }
                        }
                    }

                    // оформление границ таблицы
                    TableBorDers(ws,2, 1, currentRow, 10);

                    // сохранение файла
                    package.SaveAs(new FileInfo(filePath));
                }
            }

        }

        private void AddHeaders(ExcelWorksheet worksheet, List<string> headers)
        {
            if (headers.Count == 0) return;
            foreach (var header in headers)
            {
                if (header == null || header == "") continue;

                // пробуем расчленить на данные
                var headerData = header.Split('_');

                string name = "";
                int row = 0;
                int col = 0;
                char mergeOrientation = ' ';
                int mergeCellsCount = 0;

                    try
                    {
                       name = headerData[0];
                       row = Int16.Parse(headerData[1]);
                       col = Int16.Parse(headerData[2]);
                       mergeOrientation = Char.Parse(headerData[3]);
                       if(headerData.Count()>4) mergeCellsCount = Int16.Parse(headerData[4]);

                    }
                    catch
                    {
                        MessageBox.Show("Ошибка настройки шапки таблицы");
                        return;
                    }

                // настраиваем столбец
                var cell = worksheet.Cells[row, col];

                // объединение ячеек, если это необходимо
                if (mergeOrientation != 'n' && mergeOrientation != ' ')
                {
                    if (mergeCellsCount == 0)
                    {
                        MessageBox.Show($"Ошибка настройки шапки таблицы. Не верно указано количество объединяемых ячеек для {name}");
                        return;
                    }

                    if (mergeOrientation == 'v')    // если вертикальное объединение 
                    {
                        cell = worksheet.Cells[row, col, row + mergeCellsCount - 1, col];
                        worksheet.Cells[row, col, row + mergeCellsCount - 1, col].Merge = true;
                    }
                    else    // если горизонтальное объединение 
                    {
                        cell = worksheet.Cells[row, col, row, col + mergeCellsCount - 1];
                        worksheet.Cells[row, col, row, col + mergeCellsCount-1].Merge = true;
                    } 
                }

                // наименование столбца
                cell.Value = name;

                // ширина стобца
                if (mergeOrientation != 'h') worksheet.Column(col).Width = 20;

                // высота строки
                if (mergeOrientation != 'м') worksheet.Row(row).Height = 30;

                // оформление текста и границ
                cell.Style.WrapText = true;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell.Style.Border.BorderAround(ExcelBorderStyle.Medium);
            }
        }




        private void AddSignDataToTable(ExcelWorksheet worksheet, Stand stand, Sign sign, ref int currentRow)
        {
            currentRow += 1;
            
            AddTextToCell(currentRow, 1, (currentRow - 1).ToString()); //#
            AddTextToCell(currentRow, 2, stand.AxisName);
            AddTextToCell(currentRow, 3, sign.Number);
            AddTextToCell(currentRow, 4, sign.Name);
            AddTextToCell(currentRow, 5, sign.TypeSize);
            AddTextToCell(currentRow, 6, stand.PK);
            AddTextToCell(currentRow, 7, stand.Side);
            AddTextToCell(currentRow, 8, sign.Doubled);
            AddTextToCell(currentRow, 9, sign.Existence);

            worksheet.Row(currentRow).Height = 30;

            void AddTextToCell(int row, int column, string txt)
            {
                ExcelRange cell;
                cell = worksheet.Cells[row, column];
                cell.Value = txt;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell.Style.WrapText = true;
            }
        }


        private void  TableBorDers(ExcelWorksheet worksheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            ExcelRange cell;
            cell = worksheet.Cells[startRow, startColumn, endRow, endColumn];
            cell.Style.Border.BorderAround(ExcelBorderStyle.Medium);

            // внутрекнние границы столбцов
            for (int i = startColumn; i < endColumn; i++)
            {
                cell = worksheet.Cells[startRow, i, endRow, i];
                cell.Style.Border.Right.Style = ExcelBorderStyle.Medium;
            }

            // внутрекнние границы строк
            for (int i = startRow; i < endRow; i++)
            {
                cell = worksheet.Cells[i, startColumn, i, endColumn];
                cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }
        }





        // ************************************************************************ JSON ************************************************************************ //

        // метод сохранения в json
        private void SaveToJson<T>(T obj, string path)
        {
            string json = JsonSerializer.Serialize(obj, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }

        // метод загрузки из json
        private T LoadFromJson<T>(string path)
        {
            string json = File.ReadAllText(path);
            var obj = JsonSerializer.Deserialize<T>(json);
            return obj;
        }

        // helper для сохранения имен в шапке таблиц и 
        public void CreateTableHeader()
        {
            // name_row_column_merge(n-none, h-horizontal, v-vertical)_countOfMerge

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
            SaveToJson(signsHeader, _jsonTableHeaderSignsPath);


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
            SaveToJson(marksHeader, _jsonTableHeaderMarksPath);

            // путь к наименованием знаков
            var excelPath = Path.Combine(_settingsFolder.FullName, "Дорожные знаки.xlsx");

            // словарь для имен знаков
            Dictionary<string, string> signNames = new Dictionary<string, string>();

            // открываем excel файл и заполняем словарь
            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                var ws = package.Workbook.Worksheets[0];    // лист
                var last_row = ws.Dimension.End.Row;         // последняя строка

                for (int i = 1; i < last_row + 1; i++)
                {
                    var cur_num = ws.Cells[i, 3].Value == null ? "" : ws.Cells[i, 3].Value.ToString();
                    var cur_name = ws.Cells[i, 4].Value == null ? "" : ws.Cells[i, 4].Value.ToString();
                    if (cur_num != "" && !signNames.ContainsKey(cur_num)) signNames.Add(cur_num, cur_name);    // заполняем словарь
                }
            }
            SaveToJson<Dictionary<string, string>>(signNames, _jsonTableNamesSignsPath);
        }

         


        //    //System.Windows.MessageBox.Show("SAWD");

        //    SaveFileDialog saveFileDialog = new SaveFileDialog // создаем новое окно сохранения
        //    {
        //        Title = "Сохранить файл Excel",
        //        Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*",
        //        DefaultExt = "xlsx",
        //        FileName = "НовыйФайл.xlsx"
        //    };

        //    // показываем окно
        //    if (saveFileDialog.ShowDialog() == true)
        //    {
        //        string filePath = saveFileDialog.FileName;
        //        MessageBox.Show("Файл будет сохранён в: " + filePath);

        //        // создаём новый Excel-файл
        //        using (var workbook = new XLWorkbook())
        //        {

        //            // добавляем лист
        //            var worksheet = workbook.Worksheets.Add("Лист1");

        //            // записываем данные в ячейку A1
        //            worksheet.Cell("A1").Value = "Привет, Excel!";

        //            // сохраняем в файл
        //            workbook.SaveAs(filePath);
        //        }

        //        //тут можно сохранять Excel, например:
        //        // workbook.SaveAs(filePath);


        //    }




        //if (saveFileDialog.ShowDialog() == true)
        //{

        //    //запись пути в settings
        //    //    if (saveFileDialog.FileName != null)
        //    //    {
        //    //        string directoryPath = System.IO.Path.GetDirectoryName(saveFileDialog.FileName);
        //    //        settings.save_file_path = directoryPath;

        //    //        if (File.Exists(saveFileDialog.FileName) == false)
        //    //        {
        //    //            File.Copy(path_temp_pdf, saveFileDialog.FileName); // просто копируем файл
        //    //        }
        //    //        else
        //    //        {
        //    //            try
        //    //            {
        //    //                System.IO.File.Delete(saveFileDialog.FileName);    // сначала удаляем файл
        //    //                File.Copy(path_temp_pdf, saveFileDialog.FileName); // копируем файл
        //    //            }
        //    //            catch
        //    //            {
        //    //                MessageBox.Show("Ошибка сохранения файла PDF. Возможно он открыт в другой программе", "Ошибка сохранения", MessageBoxButton.OK, MessageBoxImage.Error);
        //    //            }
        //    //        }
        //    //    }

        //    //    pdf_save = true;

        //}



        //if (openFileDialog.ShowDialog() == true)
        //{
        //    //запись пути в settings
        //    if (openFileDialog.FileNames != null)
        //    {
        //        foreach (var file in openFileDialog.FileNames)
        //        {
        //            string directoryPath = System.IO.Path.GetDirectoryName(openFileDialog.FileName);
        //            settings.open_file_path = directoryPath;

        //            FileInfo fileInfo = new FileInfo(file);
        //            HashSum fileHash = new HashSum(file);

        //            var hashSum = Combobox_HashSum.SelectedIndex == 0 ? fileHash.GetCRC32() : fileHash.GetMD5();

        //            DateTime dateTime = DateTime.Parse(fileInfo.LastWriteTime.ToString());
        //            string resultTime = dateTime.ToString("dd.MM.yyyy H:mm"); // Используем формат без секунд

        //            FileData fileData = new FileData
        //            {
        //                File_Name = System.IO.Path.GetFileName(file),
        //                File_Sum = hashSum,
        //                File_Size = fileInfo.Length,
        //                File_Date = resultTime,

        //            };

        //            if (ListviewList_1.Any(i => i.File_Name == fileData.File_Name && i.File_Date == fileData.File_Date))
        //            {
        //                MessageBox.Show($"Файл {fileData.File_Name} от {fileData.File_Date} уже добавлен", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Exclamation);
        //            }
        //            else
        //            {
        //                ListviewList_1.Add(fileData);
        //                FileListview.Items.Refresh();
        //            }
        //        }

        //        if (ListviewList_1.Count > 1) { Combobox_form.SelectedIndex = 1; } else { Combobox_form.SelectedIndex = 0; }





    }
}
