using TSODD;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;


namespace TSODD.forms
{

    /// <summary>
    ///  Класс для заполнения DataGrid
    /// </summary>
    public class DataGridValue : INotifyPropertyChanged
    {

        public string Name { get; set; }
        public ObjectId ID { get; set; }


        private string _dataGridName;
        public string DataGridName
        {
            get => _dataGridName;
            set
            {
                if (_dataGridName == value) return;
                _dataGridName = value;
                OnPropertyChanged(nameof(DataGridName));
            }
        }

        private bool _value;
        public bool Value
        {
            get => _value;
            set
            {
                if (_value == value) return;
                _value = value;
                OnPropertyChanged(nameof(Value));
            }
        }

        private bool _isEnabled;
        public bool IsEnabled
        {
            get => _isEnabled;
            set
            {
                if (_isEnabled == value) return;
                _isEnabled = value;
                OnPropertyChanged(nameof(IsEnabled));
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    }


    /// <summary>
    /// Логика взаимодействия для LoadBlock.xaml
    /// </summary>
    public partial class LoadBlock : Window
    {

        private ObservableCollection<DataGridValue> dataGridValues = new ObservableCollection<DataGridValue>();
        private List<ObjectId> invisibleObjects = new List<ObjectId>();

        private List<string> signGroups = new List<string>();
        private List<string> markGroups = new List<string>();

        private BlockTableRecord _btr;
        private BlockReference _br;
        private bool _updateFlag = false;
        private Bitmap _bmpBlockIcon = null;
        private Extents3d _extents;


        public LoadBlock()
        {
            InitializeComponent();

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var locationX = doc.Window.DeviceIndependentLocation.X;
            var screenWidth = SystemParameters.PrimaryScreenWidth;

            this.Left = locationX + screenWidth / 2 - this.Width / 2;
            this.Top = SystemParameters.PrimaryScreenHeight / 2 - this.Height / 2;

            gridAttributes.ItemsSource = dataGridValues;

            // выбираем тип блока
            rb_Stand.IsChecked = true;
            this.Closing += LoadBlock_Closing;

            // заполняем группы
            using (var extDb = new Database(false, true))
            {
                // Открываем внешний DWG для записи
                extDb.ReadDwgFile(FilesLocation.dwgBlocksPath, FileShare.None, false, "");
                extDb.CloseInput(true);

                using (var trExt = extDb.TransactionManager.StartTransaction())
                {
                    signGroups = TsoddBlock.GetListGroups("SIGN_GROUPS", trExt, extDb);
                    markGroups = TsoddBlock.GetListGroups("MARK_GROUPS", trExt, extDb);
                }
            }
        }




        // обработчик выбора блока
        private void bt_blockSelector_Click(object sender, RoutedEventArgs e)
        {
            var listElementsID = RibbonInitializer.Instance.GetAutoCadSelectionObjectsId(new List<string> { "INSERT" }, true);
            if (listElementsID == null || listElementsID.Count != 1) return;

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {

                    _updateFlag = false;
                    RibbonInitializer.Instance.readyToDeleteEntity = false; // запрещаем анализировать объекты при удалении

                    var id = listElementsID[0];
                    _br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);

                    // если блок динамический — берем базовое определение
                    ObjectId defId = _br.IsDynamicBlock
                        ? _br.DynamicBlockTableRecord
                        : _br.BlockTableRecord;

                    // открываем определение, на которое указывает вставка
                    _btr = (BlockTableRecord)tr.GetObject(defId, OpenMode.ForRead);

                    // заполняем таблицу атрибутов
                    dataGridValues.Clear();
                    invisibleObjects.Clear();


                    // для отслеживания дубликатов
                    HashSet<string> dublicateValues = new HashSet<string> { "STAND", "ОСЬ", "SIGN", "GROUP", "STANDHANDLE", "DOUBLED",
                                                                             "НОМЕР_ЗНАКА", "TYPESIZE", "SIGNEXISTENCE", "MARK", "MATERIAL","MARKEXISTENCE" };

                    List<string> preSelect = new List<string>
                    {
                        "m","м","т","таможня","опасность","стоп","контроль",
                        "customs","danger","stop","такси","зона","объезд",
                        "ДПС","дпс","милиция","радио","связь","мгц","wc","мин"
                    };

                    foreach (ObjectId arId in _btr)
                    {
                        //var ar = (AttributeReference)tr.GetObject(arId, OpenMode.ForRead);
                        Entity ent = (Entity)tr.GetObject(arId, OpenMode.ForRead);

                        if (ent.Visible == false)   // если объекты невидимы, то не добавляем их в блок
                        {
                            invisibleObjects.Add(ent.Id);
                            continue;
                        }

                        if (ent is AttributeDefinition attr)
                        {
                            if (!dublicateValues.Contains(attr.Tag.ToString()))
                            {
                                DataGridValue value = new DataGridValue { ID = ent.Id, Name = attr.Tag, DataGridName = $"атрибут: {attr.Tag}", Value = false, IsEnabled = true };
                                if (value.Name.Contains("PK_VAL"))
                                {
                                    value.Value = true;
                                    value.IsEnabled = false;
                                }
                                dataGridValues.Add(value);
                                dublicateValues.Add(attr.Tag);
                            }
                        }

                        string txt = string.Empty;
                        bool matchValue = false;

                        switch (ent.GetRXClass().DxfName)
                        {
                            case "TEXT":
                                var txt_ent = ent as DBText;
                                txt = txt_ent.TextString;
                                break;

                            case "MTEXT":
                                var mtxt_ent = ent as MText;
                                txt = mtxt_ent.Text;
                                break;
                        }

                        if (string.IsNullOrEmpty(txt)) continue;

                        if (!dublicateValues.Contains(txt))
                        {
                            matchValue = preSelect.Any(p => p.Equals(txt, StringComparison.OrdinalIgnoreCase));
                            if (matchValue)
                            {
                                dataGridValues.Add(new DataGridValue { ID = ent.Id, Name = txt, DataGridName = $"текст: {txt}", Value = true, IsEnabled = true });
                            }
                            else
                            {
                                dataGridValues.Add(new DataGridValue { ID = ent.Id, Name = txt, DataGridName = $"текст: {txt}", Value = false, IsEnabled = true });
                            }
                            dublicateValues.Add(txt);
                        }
                    }

                    // прописываем имя блока 
                    tb_blockName.Text = _btr.Name;

                    // заполняем имя знака (если выбран соответствующий radiobutton)
                    FillSignName();

                    _updateFlag = true;

                    RefreshBlockImage(dataGridValues);  // обновляем картинку

                    this.Focus();
                    tr.Commit();
                }
            }
        }


        private void RefreshBlockImage(ObservableCollection<DataGridValue> collection = null)
        {
            if (_btr == null) return;
            if (_updateFlag == false) return;
            _bmpBlockIcon?.Dispose();

            RibbonInitializer.Instance.readyToDeleteEntity = false; // запрещаем анализировать объекты при удалении

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {

                    ObjectIdCollection idsToClone;
                    DataGridValue match;

                    // словарь для значений атрибутов
                    Dictionary<string, string> brValues = new Dictionary<string, string>();
                    foreach (ObjectId arId in _br.AttributeCollection)
                    {
                        var ar = tr.GetObject(arId, OpenMode.ForRead) as AttributeReference;
                        if (ar == null || ar.Visible == false) continue;
                        if (!brValues.ContainsKey(ar.Tag)) brValues.Add(ar.Tag, ar.TextString);
                    }

                    // копируем сущности из исходного блока (text и mtext)
                    idsToClone = new ObjectIdCollection();

                    BlockTableRecord dynDef = (BlockTableRecord)tr.GetObject(
                    _br.DynamicBlockTableRecord, OpenMode.ForRead);

                    foreach (ObjectId id in _btr)
                    {
                        Entity ent = (Entity)tr.GetObject(id, OpenMode.ForRead);

                        // определяем тип бъекта
                        switch (ent.GetRXClass().DxfName)
                        {
                            case "MTEXT":       // Мтекст
                                var txt = ent as MText;
                                if (string.IsNullOrEmpty(txt.Text)) continue;

                                match = collection.FirstOrDefault(a => a.Name.Equals(txt.Text, StringComparison.OrdinalIgnoreCase));
                                if (match == null)
                                {
                                    continue;
                                }
                                else
                                {
                                    if (match.Value == false) continue;
                                }
                                break;

                            case "TEXT":        // Текст

                                var mtxt = ent as DBText;
                                if (string.IsNullOrEmpty(mtxt.TextString)) continue;

                                match = collection.FirstOrDefault(a => a.Name.Equals(mtxt.TextString, StringComparison.OrdinalIgnoreCase));
                                if (match == null)
                                {
                                    continue;
                                }
                                else
                                {
                                    if (match.Value == false) continue;
                                }
                                break;

                            case "ATTDEF":      // Атрибут

                                var ad = ent as AttributeDefinition;
                                if (ad.Visible == false) continue;

                                match = collection.FirstOrDefault(a => a.Name.Equals(ad.Tag.ToString(), StringComparison.OrdinalIgnoreCase));
                                if (match == null)
                                {
                                    continue;
                                }
                                else
                                {
                                    if (match.Value == false) continue;
                                }
                                break;
                        }

                        if (ent is Entity) idsToClone.Add(id);
                    }

                    // временный блок для картинки
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);

                    if (bt.Has("tempBlockForPreviewIcon"))
                    {
                        var id = bt["tempBlockForPreviewIcon"];
                        BlockTableRecord oldBtr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForWrite);
                        oldBtr.Erase(true);
                    }

                    var tempBtr = new BlockTableRecord();
                    tempBtr.Name = "tempBlockForPreviewIcon"; // анонимное имя, чтобы не конфликтовать


                    ObjectId tempBtrId = bt.Add(tempBtr);
                    tr.AddNewlyCreatedDBObject(tempBtr, true);

                    var tempMap = new IdMapping();
                    db.DeepCloneObjects(idsToClone, tempBtrId, tempMap, false);

                    // настроим атрибуты для snapshot
                    foreach (var id in tempBtr)
                    {
                        Entity ent = (Entity)tr.GetObject(id, OpenMode.ForRead);
                        if (ent is AttributeDefinition attr)
                        {
                            if (brValues.TryGetValue(attr.Tag, out string value))
                            {
                                ent.UpgradeOpen();
                                attr.TextString = value;
                                attr.AdjustAlignment(db);
                            }
                        }
                    }

                    // перенесем порядок отображениия элементов из _btr в tempBtr
                    DrawOrderTable oldTable = (DrawOrderTable)tr.GetObject(_btr.DrawOrderTableId, OpenMode.ForRead);
                    ObjectIdCollection oldOrder = oldTable.GetFullDrawOrder(0);

                    TsoddBlock.SyncDrawOrder(oldOrder, tempBtr, tempMap, tr);

                    // делаем snapshot
                    _bmpBlockIcon = TsoddBlock.GetBlockBitmap(tempBtr, 250, out _extents);
                    var bmpSource = TsoddBlock.ToImageSource(_bmpBlockIcon);
                    im_blockImage.Source = bmpSource;    // обновляем картинку

                    // если это стойка, то для preview блока надо сделать маленькую картинку
                    if (rb_Stand.IsChecked == true)
                    {
                        _bmpBlockIcon = TsoddBlock.GetBlockBitmap(tempBtr, 32, out _extents);
                    }

                    // удаляем временный блок
                    tempBtr.UpgradeOpen();
                    tempBtr.Erase(true);

                    this.Focus();

                    tr.Commit();
                }
            }
            RibbonInitializer.Instance.readyToDeleteEntity = true; // разрешаем анализировать объекты при удалении
        }


        // обработчик checkbox
        private void Attr_checked(object sender, RoutedEventArgs e)
        {
            CheckBox ch = sender as CheckBox;
            var item = ch.DataContext as DataGridValue;

            if (item == null) return;

            var lv_item = dataGridValues.FirstOrDefault(i => i.Name == item.Name);
            if (lv_item != null)
            {
                string signNum = tb_blockName.Text.Trim();
                if (signNum.Contains('_')) signNum = signNum.Split('_')[0].Trim();

                // если выбран знак 
                if (rb_Sign.IsChecked == true && lv_item.Name == signNum && (bool)ch.IsChecked)
                {
                    MessageBox.Show("Для блока дорожного знака будет автоматически создан атрибут с его номером. " +
                        "Нет необходимости добавлять Mtext или Text с его номером. " +
                        "В противном случае данная информация задублируется на отображении блока");
                }

                lv_item.Value = (bool)ch.IsChecked;
            }

            RefreshBlockImage(dataGridValues);
        }


        private void rb_Stand_Checked(object sender, RoutedEventArgs e)
        {
            chb_singleSign.Visibility = System.Windows.Visibility.Collapsed;
            tb_signName.Visibility = System.Windows.Visibility.Collapsed;
            tb_signName.IsEnabled = false;

            // настраиваем группу
            cb_Groups.ItemsSource = null;
            gr_Groups.IsEnabled = false;
            bt_addGroup.Visibility = System.Windows.Visibility.Collapsed;
            RefreshBlockImage(dataGridValues);
        }

        private void rb_Sign_Checked(object sender, RoutedEventArgs e)
        {
            chb_singleSign.Visibility = System.Windows.Visibility.Visible;
            chb_singleSign.IsChecked = true;
            tb_signName.Visibility = System.Windows.Visibility.Visible;

            tb_signName.Visibility = System.Windows.Visibility.Visible;

            // заполняем имя знака
            FillSignName();

            // настраиваем группу
            gr_Groups.IsEnabled = true;
            bt_addGroup.Visibility = System.Windows.Visibility.Visible;
            cb_Groups.ItemsSource = signGroups;
            // предвыбор текущей группы
            if (TsoddHost.Current.currentSignGroup != null)
            {
                bool match = signGroups.Any(g => g == TsoddHost.Current.currentSignGroup);
                if (match) cb_Groups.SelectedValue = TsoddHost.Current.currentSignGroup;
            }
            else
            {
                if (cb_Groups.Items.Count > 0)
                {
                    cb_Groups.SelectedIndex = 0;
                    TsoddHost.Current.currentSignGroup = cb_Groups.SelectedValue.ToString();
                }
            }

            RefreshBlockImage(dataGridValues);
        }

        private void rb_Mark_Checked(object sender, RoutedEventArgs e)
        {
            chb_singleSign.Visibility = System.Windows.Visibility.Collapsed;
            tb_signName.Visibility = System.Windows.Visibility.Collapsed;
            tb_signName.IsEnabled = false;

            // настраиваем группу
            gr_Groups.IsEnabled = true;
            bt_addGroup.Visibility = System.Windows.Visibility.Visible;
            cb_Groups.ItemsSource = markGroups;
            // предвыбор текущей группы
            if (TsoddHost.Current.currentMarkGroup != null)
            {
                bool match = markGroups.Any(g => g == TsoddHost.Current.currentMarkGroup);
                if (match) cb_Groups.SelectedValue = TsoddHost.Current.currentMarkGroup;
            }
            else
            {
                if (cb_Groups.Items.Count > 0) cb_Groups.SelectedIndex = 0;
                TsoddHost.Current.currentMarkGroup = cb_Groups.SelectedValue.ToString();
            }
            RefreshBlockImage(dataGridValues);
        }

        private void chb_singleSign_Checked(object sender, RoutedEventArgs e)
        {
            FillSignName();
        }
        private void tb_blockName_TextChanged(object sender, TextChangedEventArgs e)
        {
            FillSignName();
        }


        private void FillSignName()
        {
            if (string.IsNullOrEmpty(tb_blockName.Text)) return;
            if (chb_singleSign.IsChecked == false)
            {
                tb_signName.IsEnabled = false;
            }
            if (!(bool)rb_Sign.IsChecked) return;

            string combinedName = string.Empty;
            string name_1 = tb_blockName.Text.Trim();
            string name_2 = string.Empty;
            string fullName = tb_blockName.Text.Trim();
            var signNames = JsonReader.LoadFromJson<Dictionary<string, string>>(FilesLocation.JsonTableNamesSignsPath);     // имена знаков

            // если сдвоенный знак
            if (chb_singleSign.IsChecked == false)
            {
                if (fullName.Contains("(") && fullName.Contains(")"))
                {
                    name_1 = fullName.Substring(0, fullName.IndexOf("(") - 1).Trim();

                    name_2 = fullName.Substring(1, fullName.IndexOf(")") - 1).Trim();
                    name_2 = name_2.Substring(fullName.IndexOf("(")).Trim();
                }
            }

            combinedName = GetName(name_1);
            if (!string.IsNullOrEmpty(name_2)) combinedName += $" ({GetName(name_2)})";

            // метод поиска наименования
            string GetName(string num)
            {
                string result = string.Empty;

                // отсекаем лишнее
                HashSet<char> delimetr = new HashSet<char> { '-', '_' };
                char hasDelimetr = num.FirstOrDefault(d => delimetr.Contains(d));
                if (hasDelimetr != 0) num = num.Split(hasDelimetr)[0].Trim();

                if (signNames.TryGetValue(num, out result))
                {
                    return result;
                }
                else
                {
                    return "знак не найден в БД";
                }
            }

            if (!string.IsNullOrEmpty(combinedName))
            {
                tb_signName.IsEnabled = true;
                tb_signName.Text = combinedName;
                tb_signName.Foreground = System.Windows.Media.Brushes.Black;
            }
            else
            {
                tb_signName.IsEnabled = false;
                tb_signName.Text = "Наименование знака";
                tb_signName.Foreground = System.Windows.Media.Brushes.Gray;
            }
        }


        private void bt_Ok_Click(object sender, RoutedEventArgs e)
        {

            string template = string.Empty;
            if ((bool)rb_Stand.IsChecked) template = "STAND_TEMPLATE";
            if ((bool)rb_Sign.IsChecked) template = "SIGN_TEMPLATE";
            if ((bool)rb_Mark.IsChecked) template = "MARK_TEMPLATE";

            string blockName = tb_blockName.Text.Trim(); ;
            string blockNumber = blockName;
            if (string.IsNullOrEmpty(tb_blockName.Text.Trim()))
            {
                MessageBox.Show("Ошибка. Не задано имя блока");
                return;
            }
            else
            {
                if (template == "SIGN_TEMPLATE" || template == "MARK_TEMPLATE")
                {
                    // отсекаем лишнее
                    HashSet<char> delimetr = new HashSet<char> { '-', '_' };
                    char hasDelimetr = blockName.FirstOrDefault(d => delimetr.Contains(d));
                    if (hasDelimetr != 0) blockNumber = blockName.Split(hasDelimetr)[0].Trim();
                }
            }

            // проверка на то, что объект имеет нужные атрибуты для ПК
            if ((bool)rb_Mark.IsChecked)
            {
                bool match = dataGridValues.Any(v => v.Name.Contains("PK_VAL"));
                if (!match)
                {
                    MessageBox.Show("Ошибка загрузки блока разметки. Блок должен иметь хотябы один атрибут с наименованием \"PK_VAL...\". " +
                                    "\n Пример:\"PK_VAL_1\" или \"PK_VAL_LEFT\".\n Так же, желательно, сделать данные атрибуты скрытыми.");
                    return;
                }
            }

            // проверка корректности имени, если для знака выбран тип сдвоенный
            if ((bool)rb_Sign.IsChecked && chb_singleSign.IsChecked == false)
            {
                bool match = tb_blockName.Text.Contains("(");

                MessageBox.Show("Ошибка наименования блока знака. Для сдвоенного знака необходимо указать номер второго знака в скобках." +
                                "\n Например: \" 2.4(8.1.1)\".");
                return;
            }

            // добавляем блок в БД
            TsoddBlock.AddBlockToBD(template, _btr, blockName, blockNumber, dataGridValues, invisibleObjects,
                                    _extents, _bmpBlockIcon, (bool)chb_singleSign.IsChecked, cb_Groups.SelectedValue?.ToString() ?? "");

            // заполняем список стоек 
            if ((bool)rb_Stand.IsChecked) RibbonInitializer.Instance.FillBlocksMenu(RibbonInitializer.Instance.splitStands, "STAND", blockName);

        }

        private void bt_Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void LoadBlock_Closing(object sender, CancelEventArgs e)
        {
            try
            {
                _btr?.Dispose();
                _br?.Dispose();
                _bmpBlockIcon?.Dispose();
            }
            catch { }
        }

        private void bt_addGroup_Click(object sender, RoutedEventArgs e)
        {
            GroupsAddForm groupsAddForm = new GroupsAddForm();
            groupsAddForm.ShowDialog();

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var locationX = doc.Window.DeviceIndependentLocation.X;
            var screenWidth = SystemParameters.PrimaryScreenWidth;

            // обновляем группы
            using (var extDb = new Database(false, true))
            {
                // Открываем внешний DWG для записи
                extDb.ReadDwgFile(FilesLocation.dwgBlocksPath, FileShare.None, false, "");
                extDb.CloseInput(true);

                using (var trExt = extDb.TransactionManager.StartTransaction())
                {
                    if ((bool)rb_Sign.IsChecked)
                    {
                        signGroups = TsoddBlock.GetListGroups("SIGN_GROUPS", trExt, extDb);
                        cb_Groups.ItemsSource = signGroups;
                    }
                    else
                    {
                        markGroups = TsoddBlock.GetListGroups("MARK_GROUPS", trExt, extDb);
                        cb_Groups.ItemsSource = signGroups;
                    }
                }
            }
        }


    }
}
