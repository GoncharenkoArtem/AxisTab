using Autodesk.AutoCAD.DatabaseServices;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace TSODD.Forms
{
    /// <summary>
    /// Логика взаимодействия для InsertBlockForm.xaml
    /// </summary>
    public partial class InsertBlockForm : Window
    {
        private ObservableCollection<string> _groups = new ObservableCollection<string>();
        private ObservableCollection<BlockForInsert> _blocks = new ObservableCollection<BlockForInsert>();
        private string _currentGroup = "";

        public InsertBlockForm()
        {
            InitializeComponent();

            var doc = TsoddHost.Current.doc;
            var locationX = doc.Window.DeviceIndependentLocation.X;
            var screenWidth = SystemParameters.PrimaryScreenWidth;

            this.Left = locationX + screenWidth / 2 - this.Width / 2;
            this.Top = SystemParameters.PrimaryScreenHeight / 2 - this.Height / 2;

            lv_Groups.ItemsSource = _groups;
            lv_Blocks.ItemsSource = _blocks;
            cb_TypeOfElements.SelectedIndex = TsoddHost.Current.currentInsertBlockTab;
        }

        private void cb_TypeOfElements_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            bool stands = false;
            if (cb_TypeOfElements.SelectedIndex == 0) { _currentGroup = "STAND"; stands = true; }
            if (cb_TypeOfElements.SelectedIndex == 1) _currentGroup = "SIGN";
            if (cb_TypeOfElements.SelectedIndex == 2) _currentGroup = "MARK";

            TsoddHost.Current.currentInsertBlockTab = cb_TypeOfElements.SelectedIndex;

            FillListViewByGroups(stands);
            
        }

        private void FillListViewByGroups(bool stands)
        {
            _groups.Clear();
            _blocks.Clear();

            if (stands)
            {
                lb_Groups.Visibility = System.Windows.Visibility.Collapsed;
                lv_Groups.Visibility = System.Windows.Visibility.Collapsed;

                var listOfBlocks = TsoddBlock.GetListOfBlocks(_currentGroup, null);

                foreach (var block in listOfBlocks)
                {
                    _blocks.Add(new BlockForInsert { Name = block.name, BMPsource = block.img });
                }

                return;
            }
            else
            {
                lb_Groups.Visibility = System.Windows.Visibility.Visible;
                lv_Groups.Visibility = System.Windows.Visibility.Visible;
            }

            var doc = TsoddHost.Current.doc;
            var db = doc.Database;

            List<string> groups = new List<string>();

            // заполняем группы
            using (var extDb = new Database(false, true))
            {
                // Открываем внешний DWG для чтения
                extDb.ReadDwgFile(FilesLocation.dwgBlocksPath, FileShare.None, false, "");
                extDb.CloseInput(true);

                using (var trExt = extDb.TransactionManager.StartTransaction())
                {
                    groups = TsoddBlock.GetListGroups($"{_currentGroup}_GROUPS", trExt, extDb);
                }
            }

            foreach (var group in groups) { _groups.Add(group); }

            // предвыбор группы
            if (_currentGroup == "SIGN")
            {
                if (_groups.Contains(TsoddHost.Current.currentSignGroup))
                {
                    lv_Groups.SelectedItem = TsoddHost.Current.currentSignGroup;
                }
                else
                {
                    if (lv_Groups.Items.Count > 0) lv_Groups.SelectedIndex = 0; 
                }
            }

            if (_currentGroup == "MARK")
            {
                if (_groups.Contains(TsoddHost.Current.currentMarkGroup))
                {
                    lv_Groups.SelectedItem = TsoddHost.Current.currentMarkGroup;
                }
                else
                {
                    if (lv_Groups.Items.Count > 0) lv_Groups.SelectedIndex = 0;
                }
            }
        }


        private void lv_Groups_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lv_Groups.SelectedValue == null) return;
            _blocks.Clear();
            string selectedGroup = lv_Groups.SelectedValue.ToString();
            if (_currentGroup == "SIGN")
            {
                TsoddHost.Current.currentSignGroup = selectedGroup;
            }
            else
            {
                TsoddHost.Current.currentMarkGroup = selectedGroup;
            }

            var listOfBlocks = TsoddBlock.GetListOfBlocks(_currentGroup, selectedGroup);


            foreach (var block in listOfBlocks)
            {
                _blocks.Add(new BlockForInsert { Name = block.name, BMPsource = block.img });
            }
        }


        private void lv_Blocks_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            BlockInsert();
        }


        private void BlockInsert()
        {
            var selectedVal = lv_Blocks.SelectedValue as BlockForInsert;
            if (selectedVal == null) return;

            this.Close();

            switch (cb_TypeOfElements.SelectedIndex )
            { 
                case 0:
                    TsoddBlock.InsertStandOrMarkBlock(selectedVal.Name, true);
                    break;
                case 1:
                    TsoddBlock.InsertSignBlock(selectedVal.Name);
                    break;
                case 2:
                    TsoddBlock.InsertStandOrMarkBlock(selectedVal.Name, false);
                    break;
            }

        }


        private void Button_DeleteBlock_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var blockName = (BlockForInsert)button.DataContext;

            if (blockName == null) return;

            // спрашиваем
            var message = MessageBox.Show($"Удалить блок \"{blockName.Name}\" из БД?", "Удаление блока", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (message == MessageBoxResult.Yes)
            {
                TsoddBlock.DeleteBlockFromBD(blockName.Name);

                if (cb_TypeOfElements.SelectedIndex == 0)
                {
                    _blocks.Clear();
                    var listOfBlocks = TsoddBlock.GetListOfBlocks(_currentGroup, null);

                    foreach (var block in listOfBlocks)
                    {
                        _blocks.Add(new BlockForInsert { Name = block.name, BMPsource = block.img });
                    }
                }
                else
                {
                    lv_Groups_SelectionChanged(null, null);
                }
            }
        }


        private void Button_Add_Click(object sender, RoutedEventArgs e)
        {
            BlockInsert();
        }


        private void Button_Exit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }

    public class BlockForInsert
    {
        public string Name { get; set; }
        public BitmapSource BMPsource { get; set; }
    }


}
