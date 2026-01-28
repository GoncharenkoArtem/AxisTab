using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;



namespace TSODD.forms
{



    /// <summary>
    /// Логика взаимодействия для LineTypeForm.xaml
    /// </summary>
    public partial class LineTypeForm : Window
    {

        private ObservableCollection<string> patterns = new ObservableCollection<string>();
        private ObservableCollection<ColorItem> defaultColors = new ObservableCollection<ColorItem>();
        private ObservableCollection<ColorItem> colors = new ObservableCollection<ColorItem>();
        private ObservableCollection<double> thickness = new ObservableCollection<double>();
        private ObservableCollection<double> offset = new ObservableCollection<double>();

        public ColorItem selectedColorItem;
        private ColorsGallery _colorsGallery;
        private string _currentTab;


        public LineTypeForm()
        {
            InitializeComponent();

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var locationX = doc.Window.DeviceIndependentLocation.X;
            var screenWidth = SystemParameters.PrimaryScreenWidth;

            this.Left = locationX + screenWidth / 2 - this.Width / 2;
            this.Top = SystemParameters.PrimaryScreenHeight / 2 - this.Height / 2;

            gridPattern.ItemsSource = patterns;
            gridColor.ItemsSource = colors;
            gridThickness.ItemsSource = thickness;
            gridOffset.ItemsSource = offset;

            FillDefaultColors();
            FillComboboxLineTypes();
            TabControlChanged();

            cb_TypeOfLTForDelete.SelectedIndex = 0;
            _currentTab = "Tabitem_singleLineType";
        }



        // подписка вкладок на их активацию
        private void TabControlChanged()
        {
            // Подписываемся на событие, срабатывающее на изменение свойства IsSelected у вкладки Tabitem_singleLineType
            Tabitem_singleLineType.Loaded += (s, e) =>
            {
                DependencyPropertyDescriptor.FromProperty(System.Windows.Controls.TabItem.IsSelectedProperty, typeof(System.Windows.Controls.TabItem))
                    .AddValueChanged(Tabitem_singleLineType, (s2, e2) =>
                    {
                        if (Tabitem_singleLineType.IsSelected)   // Вкладка активировалась
                        {
                            _currentTab = Tabitem_singleLineType.Name;
                            bt_AddLineType.Visibility = System.Windows.Visibility.Visible;
                        }
                    });
            };

            // Подписываемся на событие, срабатывающее на изменение свойства IsSelected у вкладки Tabitem_multiLineType
            Tabitem_multiLineType.Loaded += (s, e) =>
            {
                DependencyPropertyDescriptor.FromProperty(System.Windows.Controls.TabItem.IsSelectedProperty, typeof(System.Windows.Controls.TabItem))
                .AddValueChanged(Tabitem_multiLineType, (s2, e2) =>
                {
                    if (Tabitem_multiLineType.IsSelected)   // Вкладка активировалась
                    {
                        _currentTab = Tabitem_multiLineType.Name;
                        bt_AddLineType.Visibility = System.Windows.Visibility.Visible;
                    }
                });
            };

            // Подписываемся на событие, срабатывающее на изменение свойства IsSelected у вкладки Tabitem_deleteLineType
            Tabitem_deleteLineType.Loaded += (s, e) =>
            {
                DependencyPropertyDescriptor.FromProperty(System.Windows.Controls.TabItem.IsSelectedProperty, typeof(System.Windows.Controls.TabItem))
                .AddValueChanged(Tabitem_deleteLineType, (s2, e2) =>
                {
                    if (Tabitem_deleteLineType.IsSelected)   // Вкладка активировалась
                    {
                        _currentTab = Tabitem_deleteLineType.Name;
                        bt_AddLineType.Visibility = System.Windows.Visibility.Collapsed;
                        cb_TypeOfLTForDelete_SelectionChanged(null, null);
                    }
                });
            };
        }





        // выбор сплошного паттерна
        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            CheckBox chb = sender as CheckBox;
            if ((bool)chb.IsChecked)
            {
                tb_LineLengt.Text = "";
                tb_LineLengt.IsEnabled = false;
                bt_LineLengt.IsEnabled = false;

                tb_SpaceLengt.Text = "";
                tb_SpaceLengt.IsEnabled = false;
                bt_SpaceLengt.IsEnabled = false;

                bt_Point.IsEnabled = false;

                tb_Pattern.Text = "[continuous]";
                tb_Pattern.IsEnabled = false;
            }
            else
            {
                tb_LineLengt.IsEnabled = true;
                bt_LineLengt.IsEnabled = true;

                tb_SpaceLengt.IsEnabled = true;
                bt_SpaceLengt.IsEnabled = true;

                bt_Point.IsEnabled = true;

                tb_Pattern.Text = "";
                tb_Pattern.IsEnabled = true;
            }
        }

        // штрих
        private void bt_LineLengt_Click(object sender, RoutedEventArgs e)
        {
            double.TryParse(tb_LineLengt.Text, out double value);
            if (value != 0) tb_Pattern.Text = tb_Pattern.Text + $"[{Math.Abs(value)}]";
        }

        // пробел
        private void bt_SpaceLengt_Click(object sender, RoutedEventArgs e)
        {
            double.TryParse(tb_SpaceLengt.Text, out double value);
            if (value != 0)
            {
                if (value > 0) value = value * -1;
                tb_Pattern.Text = tb_Pattern.Text + $"[{value}]";
            }
        }

        // точка
        private void bt_Point_Click(object sender, RoutedEventArgs e)
        {
            tb_Pattern.Text = tb_Pattern.Text + $"[0]";
        }


        private void tb_Pattern_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshPatternCanvas(out _);
        }




        private bool RefreshPatternCanvas(out List<double> listValues)
        {
            CanvasLineType.Children.Clear();    // очищаем canvas

            // пробуем собрать паттерн
            DoubleCollection dashArray = new DoubleCollection();
            if (!ParsePattern(tb_Pattern.Text, out listValues)) return false;

            // проверка на ошибку
            if (!(bool)chb_Continuous.IsChecked)
            {
                if (listValues.Count < 2 || listValues.Count > 4) return false; // ошибка паттерна - выходим
            }

            // подписи
            if ((bool)chb_Continuous.IsChecked)
            {
                Label label = new Label();
                label.Content = "continuous";
                label.Foreground = Brushes.Gray;
                label.Background = Brushes.White;
                Canvas.SetLeft(label, 45);
                Canvas.SetTop(label, -15);
                CanvasLineType.Children.Add(label);
            }
            else
            {
                //заполняем dashArray с учетом масштаба
                foreach (var val in listValues)
                {
                    if (val == 0)
                    {
                        dashArray.Add(Math.Abs(0.4));
                    }
                    else
                    {
                        dashArray.Add(Math.Abs(val * 2));
                    }
                }

                double lastPos = 0;
                bool up = true;

                while (lastPos < 150)
                {
                    foreach (var val in dashArray)
                    {
                        Label label = new Label();

                        if (val != 0.4) //point
                        {
                            label.Content = Math.Abs(val / 2);
                        }

                        label.Foreground = Brushes.Gray;
                        label.Background = Brushes.White;
                        Canvas.SetLeft(label, lastPos + Math.Abs(val) - 7);

                        if (up) Canvas.SetTop(label, -13);
                        if (!up) Canvas.SetTop(label, 5);
                        CanvasLineType.Children.Add(label);

                        lastPos = lastPos + Math.Abs(val * 2.5);
                        up = !up;
                    }
                }
            }

            // добавляем линию
            Line line = new Line();
            line.Stroke = Brushes.Black;
            line.StrokeThickness = 2.5;
            line.X1 = 0;
            line.Y1 = 10;
            line.X2 = 160;
            line.Y2 = 10;
            if (!(bool)chb_Continuous.IsChecked) line.StrokeDashArray = dashArray;

            CanvasLineType.Children.Add(line);

            return true;
        }

        // парсит значения паттерна
        private bool ParsePattern(string value, out List<double> resultList)
        {
            resultList = new List<double>();

            // парсим значения textbox
            string pattern = value;
            if (string.IsNullOrEmpty(pattern)) return false;

            pattern = pattern.Replace("]", "");
            var patternsList = pattern.Split('[').ToList();

            // чистим список
            foreach (var item in patternsList) item.Trim();
            patternsList = patternsList.Where(e => e != "").ToList();

            bool lastDistSign = true;   // true - положительное

            if (patternsList.Count > 1)
            {
                for (int i = 0; i < patternsList.Count; i++)
                {
                    // пробуем получить значение 
                    double dist = double.NaN;
                    if (double.TryParse(patternsList[i], out dist))
                    {
                        if (dist < 0 && resultList.Count == 0) continue; // если первый элемент это пробел, то скипаем его
                        if ((lastDistSign == true && dist > 0 && resultList.Count > 0) ||
                            (lastDistSign == false && dist < 0 && resultList.Count > 0))
                        {
                            resultList[resultList.Count - 1] = resultList.Last() + dist; // обновляем данные
                        }
                        else
                        {
                            resultList.Add(dist);
                            lastDistSign = dist > 0 ? true : false;
                            if (dist == 0) lastDistSign = !lastDistSign;
                        }
                    }
                }
            }
            else
            {
                if (patternsList.First() == "continuous")
                {
                    resultList.Add(1000);
                    resultList.Add(0);
                }
            }
            return true;
        }


        // добавляет паттерн в таблицу
        private void bt_AddPattern_Click(object sender, RoutedEventArgs e)
        {
            if (RefreshPatternCanvas(out var list)) // если с паттерном все хорошо, то добавляем в таблицу
            {
                string value = "";

                // формируем запись 
                if (list.Count == 2 && list.First() == 1000) // значит "continuous"
                {
                    value = "[continuous]";
                }
                else
                {
                    foreach (var val in list) value = value + $"[{val}]";
                }

                if (!patterns.Any(p => p == value)) patterns.Add(value);

                tb_LineLengt.Clear();
                tb_SpaceLengt.Clear();
                tb_Pattern.Clear();

            }
        }


        // удаляет элемент с таблицы паттернов
        private void ButtonDelete_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as System.Windows.Controls.Button;
            var value = button?.DataContext;

            if (value == null) return;

            switch (value)
            {
                case var i when i is string str:
                    patterns.Remove(str);
                    break;

                case var s when s is ColorItem coli:
                    colors.Remove(coli);
                    break;

                case var d when d is double dbl:

                    if (button.Name == "thickness") thickness.Remove(dbl);
                    if (button.Name == "offset") offset.Remove(dbl);
                    break;
            }
        }

        private void FillDefaultColors()
        {
            lv_DefaultColors.Items.Clear();
            List<ColorItem> colors = new List<ColorItem>();
            System.Windows.Media.Color tempColorWpf;
            Autodesk.AutoCAD.Colors.Color tempColorAcad;
            string txt;

            // белый-черный
            colors.Add(new ColorItem
            {
                Text = "белый / черный",
                Color_1 = System.Windows.Media.Color.FromArgb(255, 0, 0, 0),
                Color_2 = System.Windows.Media.Color.FromArgb(255, 255, 255, 255),
                Index = 7
            });

            // желтый
            tempColorAcad = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2);
            tempColorWpf = System.Windows.Media.Color.FromArgb(tempColorAcad.ColorValue.A, tempColorAcad.ColorValue.R, tempColorAcad.ColorValue.G, tempColorAcad.ColorValue.B);
            txt = $"цвет 2";
            colors.Add(new ColorItem { Text = txt, Color_1 = tempColorWpf, Color_2 = tempColorWpf, Index = 2 });

            // оранжевый
            tempColorAcad = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 40);
            tempColorWpf = System.Windows.Media.Color.FromArgb(tempColorAcad.ColorValue.A, tempColorAcad.ColorValue.R, tempColorAcad.ColorValue.G, tempColorAcad.ColorValue.B);
            txt = $"цвет 40";
            colors.Add(new ColorItem { Text = txt, Color_1 = tempColorWpf, Color_2 = tempColorWpf, Index = 40 });

            // красный
            tempColorAcad = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1);
            tempColorWpf = System.Windows.Media.Color.FromArgb(tempColorAcad.ColorValue.A, tempColorAcad.ColorValue.R, tempColorAcad.ColorValue.G, tempColorAcad.ColorValue.B);
            txt = $"цвет 1";
            colors.Add(new ColorItem { Text = txt, Color_1 = tempColorWpf, Color_2 = tempColorWpf, Index = 1 });

            // синий
            tempColorAcad = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 5);
            tempColorWpf = System.Windows.Media.Color.FromArgb(tempColorAcad.ColorValue.A, tempColorAcad.ColorValue.R, tempColorAcad.ColorValue.G, tempColorAcad.ColorValue.B);
            txt = $"цвет 5";
            colors.Add(new ColorItem { Text = txt, Color_1 = tempColorWpf, Color_2 = tempColorWpf, Index = 5 });

            // зеленый
            tempColorAcad = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 82);
            tempColorWpf = System.Windows.Media.Color.FromArgb(tempColorAcad.ColorValue.A, tempColorAcad.ColorValue.R, tempColorAcad.ColorValue.G, tempColorAcad.ColorValue.B);
            txt = $"цвет 82";
            colors.Add(new ColorItem { Text = txt, Color_1 = tempColorWpf, Color_2 = tempColorWpf, Index = 82 });

            lv_DefaultColors.ItemsSource = colors;
        }


        private void FillComboboxLineTypes()
        {
            string lineType_1 = cb_firstLineType.SelectedIndex > -1 ? cb_firstLineType.SelectedValue.ToString() : "";
            string lineType_2 = cb_secondLineType.SelectedIndex > -1 ? cb_secondLineType.SelectedValue.ToString() : "";

            cb_firstLineType.Items.Clear();
            cb_secondLineType.Items.Clear();
            var listOfLineTypes = LineTypeReader.Parse();

            // заполняем заново
            foreach (var lineType in listOfLineTypes)
            {
                if (lineType.Type == AcadDefType.LineType)
                {
                    cb_firstLineType.Items.Add(lineType.Name);
                    cb_secondLineType.Items.Add(lineType.Name);
                }
            }

            // выбираем то, что было выбрано 
            if (listOfLineTypes.Any(lt => lt.Name == lineType_1))
            {
                cb_firstLineType.SelectedValue = lineType_1;
                cb_LineType_SelectionChanged(cb_firstLineType, null);
            }
            if (listOfLineTypes.Any(lt => lt.Name == lineType_2))
            {
                cb_secondLineType.SelectedValue = lineType_2;
                cb_LineType_SelectionChanged(cb_secondLineType, null);
            }
        }

        private void lv_DefaultColors_Selected(object sender, RoutedEventArgs e)
        {
            selectedColorItem = lv_DefaultColors.SelectedItem as ColorItem;
            if (selectedColorItem == null) return;
            brd_ColorBorder_1.Background = new SolidColorBrush(selectedColorItem.Color_1);
            brd_ColorBorder_2.Background = new SolidColorBrush(selectedColorItem.Color_2);
        }

        private void bt_AddColor_Click(object sender, RoutedEventArgs e)
        {
            bool match = colors.Any(c => c.Index == selectedColorItem.Index);
            if (!match) colors.Add(selectedColorItem);
        }

        private void bt_ColorsGallery_Click(object sender, RoutedEventArgs e)
        {
            _colorsGallery = new ColorsGallery(this);
            _colorsGallery.Show();
        }

        private void bt_addThickness_Click(object sender, RoutedEventArgs e)
        {
            double value = 0;
            double.TryParse(tb_Thickness.Text, out value);
            if (value != 0)
            {
                bool match = thickness.Any(x => x == value);
                if (!match) thickness.Add(value);
            }
        }

        private void bt_addOffset_Click(object sender, RoutedEventArgs e)
        {
            double value = 0;
            double.TryParse(tb_Offset.Text, out value);
            if (value != 0)
            {
                bool match = offset.Any(x => x == value);
                if (!match) offset.Add(value);
            }
        }


        private void DeleteLineType(object sender, RoutedEventArgs e)
        {
            var bt = sender as Button;
            var dc = bt.DataContext;
            dynamic dynamicObj = dc;
            string name = dynamicObj.Column1;

            if (!string.IsNullOrEmpty(name))
            {
                LineTypeReader.DeleteLineTypeFromBDs(name);
            }
            cb_TypeOfLTForDelete_SelectionChanged(null, null);
        }


        private void cb_LineType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cb = sender as ComboBox;

            TextBlock textBlockIsForMultiline;
            ItemsControl itemsControlPattern;
            ItemsControl itemsControlColor;
            TextBlock textBlockThickness;

            if (cb.Name == "cb_firstLineType")
            {
                textBlockIsForMultiline = tb_firstLtIsForMultiline;
                itemsControlPattern = ic_firstLtPatterns;
                itemsControlColor = ic_firstLtColors;
                textBlockThickness = tb_firstLtThickness;
            }
            else
            {
                textBlockIsForMultiline = tb_secondLtIsForMultiline;
                itemsControlPattern = ic_secondLtPatterns;
                itemsControlColor = ic_secondLtColors;
                textBlockThickness = tb_secondLtThickness;
            }

            textBlockIsForMultiline.Text = "";
            itemsControlPattern.Items.Clear();
            itemsControlColor.Items.Clear();
            textBlockThickness.Text = "";


            if (cb.SelectedIndex < 0 || cb.SelectedValue == null) return;
            var listOfLineTypes = LineTypeReader.Parse();

            AcadLineType selectedLT = listOfLineTypes.FirstOrDefault(lt => lt.Name == cb.SelectedValue.ToString()) as AcadLineType;
            if (selectedLT == null) return;

            // только для мультилинии?
            textBlockIsForMultiline.Text = selectedLT.IsMlineElement ? "(только для двойной линии)" : "";

            // паттерны
            foreach (var patternList in selectedLT.PatternValues)
            {
                string txt_pattern = "";
                if (patternList.Count == 2 && patternList.First() == 1000)
                {
                    txt_pattern = "[continuous]";
                    itemsControlPattern.Items.Add(txt_pattern);
                }
                else
                {
                    if (patternList.Count > 1)
                    {
                        foreach (var val in patternList) txt_pattern = txt_pattern + $"[{val}]";
                        itemsControlPattern.Items.Add(txt_pattern);
                    }
                }
            }

            // цвета
            foreach (var color in selectedLT.ColorIndex)
            {
                ColorItem colorItem = new ColorItem();
                if (color == 7) // черный белый
                {
                    colorItem.Color_1 = System.Windows.Media.Color.FromArgb(255, 0, 0, 0);
                    colorItem.Color_2 = System.Windows.Media.Color.FromArgb(255, 255, 255, 255);

                }
                else
                {
                    var tempColorAcad = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, color);
                    var tempColorWpf = System.Windows.Media.Color.FromArgb(tempColorAcad.ColorValue.A, tempColorAcad.ColorValue.R, tempColorAcad.ColorValue.G, tempColorAcad.ColorValue.B);
                    colorItem.Color_1 = tempColorWpf;
                    colorItem.Color_2 = tempColorWpf;
                }
                itemsControlColor.Items.Add(colorItem);
            }

            // толщины
            string txt_width = "";
            foreach (var width in selectedLT.Width)
            {
                if (width > 0) txt_width = txt_width + $"; {width}";
            }

            if (selectedLT.Width.Count > 0) txt_width = txt_width.Substring(1).Trim();
            textBlockThickness.Text = txt_width;

        }

        private void cb_TypeOfLTForDelete_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            gridDeleteLineType.Items.Clear();
            var listOfLineTypes = LineTypeReader.Parse();
            string txt = "";

            switch (cb_TypeOfLTForDelete.SelectedIndex)
            {
                case 0: // все
                    foreach (var lt in listOfLineTypes)
                    {
                        if (lt is AcadLineType al)
                        {
                            txt = "одинарная линия";
                            if (al.IsMlineElement) txt = "только для двойной линии";
                            gridDeleteLineType.Items.Add(new { Column1 = al.Name, Column2 = txt });
                        }
                        if (lt is AcadMLineType aml)
                        {
                            txt = "двойная линия";
                            gridDeleteLineType.Items.Add(new { Column1 = aml.Name, Column2 = txt });
                        }
                    }
                    break;

                case 1: // одинарные лини
                    foreach (var lt in listOfLineTypes)
                    {
                        if (lt is AcadLineType al)
                        {
                            txt = "одинарная линия";
                            if (al.IsMlineElement) txt = "только для двойной линии";
                            gridDeleteLineType.Items.Add(new { Column1 = al.Name, Column2 = txt });
                        }
                    }
                    break;

                case 2: // двойные типы линий
                    foreach (var lt in listOfLineTypes)
                    {
                        if (lt is AcadMLineType aml)
                        {
                            txt = "двойная линия";
                            gridDeleteLineType.Items.Add(new { Column1 = aml.Name, Column2 = txt });
                        }
                    }
                    break;

                case 3: // только для двойной линии
                    foreach (var lt in listOfLineTypes)
                    {
                        if (lt is AcadLineType al)
                        {
                            if (al.IsMlineElement)
                            {
                                txt = "только для двойной линии";
                                gridDeleteLineType.Items.Add(new { Column1 = al.Name, Column2 = txt });
                            }
                        }
                    }
                    break;
            }
        }



        private void ButtonAddLineType_Click(object sender, RoutedEventArgs e)
        {
            IAcadDef lineType = null;

            if (_currentTab == "Tabitem_singleLineType") lineType = new AcadLineType(); // добавляем простой тип линии
            if (_currentTab == "Tabitem_multiLineType") lineType = new AcadMLineType(); // добавляем двойной тип линии

            if (lineType is AcadLineType al)
            {
                al.Name = tb_LineTypeNameSingle.Text;

                al.Description = al.Name;
                if ((bool)chb_ForMultiline.IsChecked)
                {
                    al.Description = "[multiLine]";
                    al.IsMlineElement = true;
                }
                bool patternIsOk = true;
                foreach (var val in patterns)
                {
                    List<double> patternList = new List<double>();
                    if (!ParsePattern(val, out patternList)) patternIsOk = false;
                    al.PatternValues.Add(patternList);
                }

                if (!patternIsOk)
                {
                    MessageBox.Show("Ошибка паттерна нового типа линии.");
                    return;
                }

                foreach (var color in colors) al.ColorIndex.Add(color.Index);

                foreach (var width in thickness) al.Width.Add(width);
            }

            if (lineType is AcadMLineType aml)
            {
                var listOfLineTypes = LineTypeReader.Parse();

                aml.Name = tb_LineTypeNameMulti.Text;
                aml.Description = aml.Name;

                string lineType_1 = cb_firstLineType.SelectedIndex > -1 ? cb_firstLineType.SelectedValue.ToString() : "";
                string lineType_2 = cb_secondLineType.SelectedIndex > -1 ? cb_secondLineType.SelectedValue.ToString() : "";

                var acadLineType_1 = listOfLineTypes.FirstOrDefault(lt => lt.Name == lineType_1);
                var acadLineType_2 = listOfLineTypes.FirstOrDefault(lt => lt.Name == lineType_2);

                if (acadLineType_1 != null && acadLineType_2 != null)
                {
                    aml.MLineLineTypes.Add((AcadLineType)acadLineType_1);
                    aml.MLineLineTypes.Add((AcadLineType)acadLineType_2);
                }

                foreach (var value in offset) aml.Offset.Add(value);
            }

            LineTypeReader.AddLineTypeToBD(lineType, messageOK: true);
            FillComboboxLineTypes();
        }

        private void ButtonExit_Click(object sender, RoutedEventArgs e)
        {
            if (_colorsGallery != null) _colorsGallery.Close();
            this.Close();
        }


    }






}
