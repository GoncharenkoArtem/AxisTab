using Autodesk.AutoCAD.DatabaseServices;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace TSODD.forms
{
    /// <summary>
    /// Логика взаимодействия для GroupsAddForm.xaml
    /// </summary>
    public partial class GroupsAddForm : Window
    {

        private ObservableCollection<string> _groups = new ObservableCollection<string>();
        private string _template = "";

        public GroupsAddForm()
        {
            InitializeComponent();

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var locationX = doc.Window.DeviceIndependentLocation.X;
            var screenWidth = SystemParameters.PrimaryScreenWidth;

            this.Left = locationX + screenWidth / 2 - this.Width / 2;
            this.Top = SystemParameters.PrimaryScreenHeight / 2 - this.Height / 2;

            gridGroups.ItemsSource = _groups;

            // выбираем знаки
            cb_TypeOfElements.SelectedIndex = 0;
        }

        private void cb_TypeOfElements_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cb_TypeOfElements.SelectedIndex == 0) _template = "SIGN_GROUPS";
            if (cb_TypeOfElements.SelectedIndex == 1) _template = "MARK_GROUPS";

            FillGridByGroups();
        }

        private void FillGridByGroups()
        {
            _groups.Clear();
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            List<string> groups = new List<string>();

            // заполняем группы
            using (var extDb = new Database(false, true))
            {
                // Открываем внешний DWG для записи
                extDb.ReadDwgFile(FilesLocation.dwgBlocksPath, FileShare.None, false, "");
                extDb.CloseInput(true);

                using (var trExt = extDb.TransactionManager.StartTransaction())
                {
                    groups = TsoddBlock.GetListGroups(_template, trExt, extDb);
                }
            }

            foreach (var group in groups) { _groups.Add(group); }
        }


        private void DeleteGroup(object sender, RoutedEventArgs e)
        {
            Button bt = sender as Button;
            string name = bt.DataContext.ToString();

            if (string.IsNullOrEmpty(name)) return;

            TsoddBlock.DeleteGroupFromBD(_template, name);

            FillGridByGroups();
        }



        private void bt_AddGroupName_Click(object sender, RoutedEventArgs e)
        {
            string name = tb_GroupName.Text;
            if (string.IsNullOrEmpty(name)) return;

            TsoddBlock.AddGroupToBD(_template, name);
            FillGridByGroups();
            tb_GroupName.Text = "";
        }



        private void bt_Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }
}
