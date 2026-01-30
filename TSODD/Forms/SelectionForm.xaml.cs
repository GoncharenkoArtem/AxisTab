
using TSODD;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace TSODD.Forms
{
    /// <summary>
    /// Логика взаимодействия для SelectionFormBlocks.xaml
    /// </summary>
    public partial class SelectionFormBlocks : Window
    {
        public int _type;
        public Dictionary<ObjectId, string> _dictionary;
        ObservableCollection<string> signsTypeList = new ObservableCollection<string>();
        ObservableCollection<string> signsExistenceList = new ObservableCollection<string>();
        ObservableCollection<string> marksMaterialList = new ObservableCollection<string>();
        ObservableCollection<string> marksExistenceList = new ObservableCollection<string>();

        private bool _isLoaded = false;

        public SelectionFormBlocks(int type, Dictionary<ObjectId, string> dictionary)
        {
            InitializeComponent();
            _type = type;
            _dictionary = dictionary;

            cb_signsType.ItemsSource = signsTypeList;
            cb_signsExistence.ItemsSource = signsExistenceList;
            cb_marksMaterial.ItemsSource = marksMaterialList;
            cb_marksExistence.ItemsSource = marksExistenceList;

            this.Closed += SelectionForm_Closed;

            RebuildForm(_type, _dictionary);

            // считаем позицию
            System.Drawing.Point cursorPos = System.Windows.Forms.Cursor.Position;
            this.Left = cursorPos.X - 350;
            this.Top = cursorPos.Y - 100;

        }

        private void SelectionForm_Closed(object sender, EventArgs e)
        {
            RibbonInitializer.Instance.selectioForm = null;
        }

        public void RebuildForm(int type, Dictionary<ObjectId, string> dictionary)
        {
            _isLoaded = false;

            bool dictionariesAreEqual = _dictionary.Count == dictionary.Count && _dictionary.OrderBy(k => k.Key).SequenceEqual(dictionary.OrderBy(k => k.Key));
            if (_type != type || !dictionariesAreEqual)
            {
                // включаем все элементы
                lb_signsType.Visibility = System.Windows.Visibility.Visible;
                lb_signsExistence.Visibility = System.Windows.Visibility.Visible;
                cb_signsType.Visibility = System.Windows.Visibility.Visible;
                cb_signsExistence.Visibility = System.Windows.Visibility.Visible;
                tb_signsType.Visibility = System.Windows.Visibility.Visible;
                tb_signsExistence.Visibility = System.Windows.Visibility.Visible;

                lb_marksMaterial.Visibility = System.Windows.Visibility.Visible;
                lb_marksExistence.Visibility = System.Windows.Visibility.Visible;
                cb_marksMaterial.Visibility = System.Windows.Visibility.Visible;
                cb_marksExistence.Visibility = System.Windows.Visibility.Visible;

                tb_marksMaterial.Visibility = System.Windows.Visibility.Visible;
                tb_marksExistence.Visibility = System.Windows.Visibility.Visible;
                this.Height = 150;
            }

            // переопределяем переменные 
            _type = type;
            _dictionary = dictionary;

            if (_type == 0 || _type == 2)  // если выбраны знаки или все элементы
            {
                if (_type == 0)
                {
                    // отключаем лишнее
                    lb_marksMaterial.Visibility = System.Windows.Visibility.Collapsed;
                    lb_marksExistence.Visibility = System.Windows.Visibility.Collapsed;
                    cb_marksMaterial.Visibility = System.Windows.Visibility.Collapsed;
                    cb_marksExistence.Visibility = System.Windows.Visibility.Collapsed;
                    tb_marksMaterial.Visibility = System.Windows.Visibility.Collapsed;
                    tb_marksExistence.Visibility = System.Windows.Visibility.Collapsed;
                    this.Height = 100;
                }

                List<ObjectId> signsListId = _dictionary.Select(s => s.Key).ToList();
                var signsType = GetAttrValues(signsListId, "TYPESIZE");
                var signsExistence = GetAttrValues(signsListId, "SIGNEXISTENCE");

                // формируем начальный список комбобокс
                signsTypeList.Clear();
                signsTypeList.Add("I"); signsTypeList.Add("II"); signsTypeList.Add("III"); signsTypeList.Add("IV");

                // формируем начальный список комбобокс
                signsExistenceList.Clear();
                signsExistenceList.Add("Установить"); signsExistenceList.Add("Демонтировать");

                // добавляем в список  новые элементы, если такие есть
                foreach (var selectedType in signsType)
                {
                    if (!signsTypeList.Contains(selectedType)) signsTypeList.Add(selectedType);
                }
                if (!signsTypeList.Contains("другое...")) signsTypeList.Add("другое...");

                foreach (var selectedExistence in signsExistence)
                {
                    if (!signsExistenceList.Contains(selectedExistence)) signsExistenceList.Add(selectedExistence);
                }
                if (!signsExistenceList.Contains("другое...")) signsExistenceList.Add("другое...");

                // предвыбор типа знака
                if (signsType.Count == 1)
                {
                    tb_signsType.Visibility = System.Windows.Visibility.Collapsed;
                    cb_signsType.Text = signsType.First();
                }

                // предвыбор наличия знака
                if (signsExistence.Count == 1)
                {
                    tb_signsExistence.Visibility = System.Windows.Visibility.Collapsed;
                    cb_signsExistence.Text = signsExistence.First();
                }
            }

            if (_type == 1 || _type == 2)  // если выбрана разметка или все элементы
            {
                if (_type == 1)
                {
                    // отключаем лишнее
                    lb_signsType.Visibility = System.Windows.Visibility.Collapsed;
                    lb_signsExistence.Visibility = System.Windows.Visibility.Collapsed;
                    cb_signsType.Visibility = System.Windows.Visibility.Collapsed;
                    cb_signsExistence.Visibility = System.Windows.Visibility.Collapsed;
                    tb_signsType.Visibility = System.Windows.Visibility.Collapsed;
                    tb_signsExistence.Visibility = System.Windows.Visibility.Collapsed;
                    this.Height = 100;
                }

                List<ObjectId> signsListId = _dictionary.Select(s => s.Key).ToList();
                var marksMaterial = GetAttrValues(signsListId, "MATERIAL");
                var marksExistence = GetAttrValues(signsListId, "MARKEXISTENCE");

                // формируем начальный список комбобокс
                marksMaterialList.Clear();
                marksMaterialList.Add("Холодный пластик"); marksMaterialList.Add("Термопластик"); marksMaterialList.Add("Краска");

                // формируем начальный список комбобокс
                marksExistenceList.Clear();
                marksExistenceList.Add("Нанести"); marksExistenceList.Add("Демаркировать");


                // добавляем в список  новые элементы, если такие есть
                foreach (var selectedMaterial in marksMaterial)
                {
                    if (!marksMaterialList.Contains(selectedMaterial)) marksMaterialList.Add(selectedMaterial);
                }
                if (!marksMaterialList.Contains("другое...")) marksMaterialList.Add("другое...");

                foreach (var selectedExistence in marksExistence)
                {
                    if (!marksExistenceList.Contains(selectedExistence)) marksExistenceList.Add(selectedExistence);
                }
                if (!marksExistenceList.Contains("другое...")) marksExistenceList.Add("другое...");

                // предвыбор материала разметки
                if (marksMaterial.Count == 1)
                {
                    tb_marksMaterial.Visibility = System.Windows.Visibility.Collapsed;
                    cb_marksMaterial.Text = marksMaterial.First();
                }

                // предвыбор наличия разметки
                if (marksExistence.Count == 1)
                {
                    tb_marksExistence.Visibility = System.Windows.Visibility.Collapsed;
                    cb_marksExistence.Text = marksExistence.First();
                }
            }
            _isLoaded = true;
        }


        private HashSet<string> GetAttrValues(List<ObjectId> listId, string attrName)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            HashSet<string> result = new HashSet<string>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var id in listId)
                {
                    if (!id.IsValid) continue;

                    DBObject dbo = (DBObject)tr.GetObject(id, OpenMode.ForRead);
                    if (dbo is BlockReference br)
                    {
                        foreach (ObjectId attr in br.AttributeCollection)
                        {
                            var at = tr.GetObject(attr, OpenMode.ForRead) as AttributeReference;
                            if (at.Tag.Equals(attrName, StringComparison.OrdinalIgnoreCase))
                            {
                                if (at.TextString.Equals("другое...")) continue;
                                result.Add(at.TextString); break;
                            }
                        }
                    }

                    if (dbo is Autodesk.AutoCAD.DatabaseServices.Polyline pl)
                    {
                        TsoddXdataElement txde = new TsoddXdataElement();
                        txde.Parse(id);
                        if (attrName == "MATERIAL") { result.Add(txde.Material); }
                        if (attrName == "MARKEXISTENCE") { result.Add(txde.Existence); }
                    }
                }
            }
            return result;
        }


        private void cb_signsType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string value = null;
            if (cb_signsType.SelectedItem != null) value = cb_signsType.SelectedItem.ToString();

            if (!_isLoaded) return; // форма еще не загрузилась
            if (cb_signsType.SelectedIndex == -1) return;   // ничего не выбрано
            if (cb_signsType.SelectedIndex == cb_signsType.Items.Count - 1)

            {  // выбран последний эдемент
                this.Hide();
                value = PromptForNotDefaultValue("Введите типоразмер знака");
                this.Show();
            }

            if (string.IsNullOrEmpty(value))
            {
                if (cb_signsType.Items.Count > 0)
                {
                    value = (string)cb_signsType.Items[0];
                    cb_signsType.SelectedIndex = 0;
                    ApplyAttributeValue(value, "TYPESIZE", "SIGN");
                    RebuildForm(_type, _dictionary);
                }
                return;        // выбрано пустое значение 
            }
            ApplyAttributeValue(value, "TYPESIZE", "SIGN");
            RebuildForm(_type, _dictionary);
        }



        private void cb_signsExistence_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string value = null;
            if (cb_signsExistence.SelectedItem != null) value = cb_signsExistence.SelectedItem.ToString();

            if (!_isLoaded) return; // форма еще не загрузилась
            if (cb_signsExistence.SelectedIndex == -1) return;   // ничего не выбрано
            if (cb_signsExistence.SelectedIndex == cb_signsExistence.Items.Count - 1)

            {  // выбран последний эдемент
                this.Hide();
                value = PromptForNotDefaultValue("Введите наличие знака");
                this.Show();
            }

            if (string.IsNullOrEmpty(value))
            {
                if (cb_signsExistence.Items.Count > 0)
                {
                    value = (string)cb_signsExistence.Items[0];
                    cb_signsExistence.SelectedIndex = 0;
                    ApplyAttributeValue(value, "SIGNEXISTENCE", "SIGN");
                    RebuildForm(_type, _dictionary);
                }
                return;        // выбрано пустое значение 
            }
            ApplyAttributeValue(value, "SIGNEXISTENCE", "SIGN");
            RebuildForm(_type, _dictionary);
        }









        private void cb_marksMaterial_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string value = null;
            if (cb_marksMaterial.SelectedItem != null) value = cb_marksMaterial.SelectedItem.ToString();

            if (!_isLoaded) return; // форма еще не загрузилась
            if (cb_marksMaterial.SelectedIndex == -1) return;   // ничего не выбрано
            if (cb_marksMaterial.SelectedIndex == cb_marksMaterial.Items.Count - 1)

            {  // выбран последний эдемент
                this.Hide();
                value = PromptForNotDefaultValue("Введите материал разметки");
                this.Show();
            }

            if (string.IsNullOrEmpty(value))
            {
                if (cb_marksMaterial.Items.Count > 0)
                {
                    value = (string)cb_marksMaterial.Items[0];
                    cb_marksMaterial.SelectedIndex = 0;
                    ApplyAttributeValue(value, "MATERIAL", "MARK");
                    RebuildForm(_type, _dictionary);
                }
                return;        // выбрано пустое значение 
            }


            ApplyAttributeValue(value, "MATERIAL", "MARK");
            RebuildForm(_type, _dictionary);
        }


        private void cb_marksExistence_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string value = null;
            if (cb_marksExistence.SelectedItem != null) value = cb_marksExistence.SelectedItem.ToString();

            if (!_isLoaded) return; // форма еще не загрузилась
            if (cb_marksExistence.SelectedIndex == -1) return;   // ничего не выбрано
            if (cb_marksExistence.SelectedIndex == cb_marksExistence.Items.Count - 1)

            {  // выбран последний эдемент
                this.Hide();
                value = PromptForNotDefaultValue("Введите наличие разметки");
                this.Show();
            }

            if (string.IsNullOrEmpty(value))
            {
                if (cb_marksExistence.Items.Count > 0)
                {
                    value = (string)cb_marksExistence.Items[0];
                    cb_marksExistence.SelectedIndex = 0;
                    ApplyAttributeValue(value, "MARKEXISTENCE", "MARK");
                    RebuildForm(_type, _dictionary);
                }
                return;        // выбрано пустое значение 
            }
            ApplyAttributeValue(value, "MARKEXISTENCE", "MARK");
            RebuildForm(_type, _dictionary);
        }



        private void ApplyAttributeValue(string value, string attributeTag, string blockName)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            // отбираем только нужные типы (знак или разметка)
            var separatedList = _dictionary.Where(i => i.Value == blockName).ToList();

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var item in separatedList)
                {
                    DBObject dbo = (DBObject)tr.GetObject(item.Key, OpenMode.ForWrite);
                    if (dbo is BlockReference br) TsoddBlock.ChangeAttribute(tr, br, attributeTag, value); // если это блок то меняем у него атрибуты
                    if (dbo is Autodesk.AutoCAD.DatabaseServices.Polyline pl)                              // если это линия, то меняем xdata
                    {
                        // получааем xData полилинии
                        var listXdata = AutocadXData.ReadXData(item.Key);

                        if (attributeTag.Equals("Material", StringComparison.InvariantCultureIgnoreCase))   // корректируем значения xData
                        { listXdata[5] = (listXdata[5].Item1, value); }
                        else { listXdata[6] = (listXdata[6].Item1, value); }

                        // сохраняем xData
                        AutocadXData.UpdateXData(item.Key, listXdata);
                    }
                }
                tr.Commit();
            }
        }

        private string PromptForNotDefaultValue(string txt)
        {
            var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            PromptStringOptions pso = new PromptStringOptions($"\n {txt} (Esc - выход): ");
            pso.AllowSpaces = true;

            var per = ed.GetString(pso);
            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n Ошибка ввода ...");
                return null;
            }
            // новое наименование группы  
            return per.StringResult;
        }


    }
}
