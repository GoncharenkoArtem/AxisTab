
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ACAD = Autodesk.AutoCAD.ApplicationServices;

namespace TSODD
{
    public class ExportExcel
    {
        EPPlusLicense _license;

        public ExportExcel()
        {
            _license = new EPPlusLicense();
            _license.SetNonCommercialOrganization("abc");
        }


        // метод эксопрта адресной ведомости знаков  эксель
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
                        if (fs != null) fs.Dispose();
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
                    var signHeaders = JsonReader.LoadFromJson<List<string>>(FilesLocation.JsonTableHeaderSignsPath);                // для шапки таблицы
                    var signNames = JsonReader.LoadFromJson<Dictionary<string, string>>(FilesLocation.JsonTableNamesSignsPath);     // имена знаков

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
                    if (currentRow > 1) TableBorders(ws, 2, 1, currentRow, 10);

                    // сохранение файла
                    package.SaveAs(new FileInfo(filePath));
                }

                ExportEnd exportEnd = new ExportEnd();
                exportEnd.ShowDialog();
            }
        }

        // метод эксопрта адресной ведомости разметки эксель
        public void ExportMarks()
        {
            var exportFileName = $"Адресная ведомость дорожной разметки {ACAD.Application.DocumentManager.MdiActiveDocument.IsNamedDrawing}";

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
                        if (fs != null) fs.Dispose();
                    }
                }

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    // что бы там не было в файле - просто очищаем его в 0
                    while (package.Workbook.Worksheets.Count > 0)
                        package.Workbook.Worksheets.Delete(0);

                    // новый лист для ведомости
                    var ws = package.Workbook.Worksheets.Add("Дорожная разметка");

                    // выгружаем данные из JSON
                    var markHeaders = JsonReader.LoadFromJson<List<string>>(FilesLocation.JsonTableHeaderMarksPath);                // для шапки таблицы

                    // шапка таблицы
                    AddHeaders(ws, markHeaders);
                    int currentRow = 2;

                    List<Mark> marksLineTypes = TsoddBlock.GetListOfLineTypes();             // список линий разметки 
                    List<Mark> marksBlocks = TsoddBlock.GetListOfRefBlocks<Mark>();          // список блоков разметки

                    List<Mark> markList = marksLineTypes.Union(marksBlocks).ToList();        // объединяем списки

                    // сортировка разметки по Осям
                    var markListGroups = markList.GroupBy(g => g.AxisName).ToList();

                    foreach (var group in markListGroups)
                    {
                        // сортировка по расстоянию
                        var orderedMarks = group.OrderBy(s => s.Distance);
                        foreach (Mark mark in orderedMarks) AddMarkDataToTable(ws, mark, ref currentRow);
                    }

                    // оформление границ таблицы
                    if (currentRow > 1) TableBorders(ws, 2, 1, currentRow, 10);

                    // сохранение файла
                    package.SaveAs(new FileInfo(filePath));
                }

                ExportEnd exportEnd = new ExportEnd();
                exportEnd.ShowDialog();
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
                    if (headerData.Count() > 4) mergeCellsCount = Int16.Parse(headerData[4]);

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
                        worksheet.Cells[row, col, row, col + mergeCellsCount - 1].Merge = true;
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

        private void AddMarkDataToTable(ExcelWorksheet worksheet, Mark mark, ref int currentRow)
        {
            currentRow += 1;

            AddTextToCell(currentRow, 1, (currentRow - 2).ToString()); //#
            AddTextToCell(currentRow, 2, mark.AxisName);
            AddTextToCell(currentRow, 3, mark.Number);
            AddTextToCell(currentRow, 4, mark.Quantity);
            AddTextToCell(currentRow, 5, mark.Square.ToString());
            AddTextToCell(currentRow, 6, mark.PK_start);
            AddTextToCell(currentRow, 7, mark.PK_end);
            AddTextToCell(currentRow, 8, mark.Side);
            AddTextToCell(currentRow, 9, mark.Material);
            AddTextToCell(currentRow, 10, mark.Existence);

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

        private void TableBorders(ExcelWorksheet worksheet, int startRow, int startColumn, int endRow, int endColumn)
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
    }
}
