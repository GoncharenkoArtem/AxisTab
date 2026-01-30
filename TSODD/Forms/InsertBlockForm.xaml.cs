using Autodesk.AutoCAD.DatabaseServices;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
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
            cb_TypeOfElements.SelectedIndex = 0;
        }

        private void cb_TypeOfElements_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cb_TypeOfElements.SelectedIndex == 0) _currentGroup = "SIGN";
            if (cb_TypeOfElements.SelectedIndex == 1) _currentGroup = "MARK";
            FillListViewByGroups();
            if (lv_Groups.Items.Count > 0) lv_Groups.SelectedIndex = 0;
        }

        private void FillListViewByGroups()
        {
            _groups.Clear();
            _blocks.Clear();
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
            if (lv_Groups.SelectedValue == null) return;

            this.Close();

            if (cb_TypeOfElements.SelectedIndex == 0)
            {
                TsoddBlock.InsertSignBlock(selectedVal.Name);
            }
            else
            {
                TsoddBlock.InsertStandOrMarkBlock(selectedVal.Name, false);
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
                lv_Groups_SelectionChanged(null, null);
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
