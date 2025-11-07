using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace TSODD
{
    /// <summary>
    /// Логика взаимодействия для OptionsForm.xaml
    /// </summary>
    public partial class OptionsForm : Window
    {
        private Options userOptions;
        private ObservableCollection<string> textStyleNames = new ObservableCollection<string>();
        private ObservableCollection<string> mleaderStyleNames = new ObservableCollection<string>();

        public OptionsForm()
        {
            InitializeComponent();

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var locationX = doc.Window.DeviceIndependentLocation.X;
            var screenWidth = SystemParameters.PrimaryScreenWidth;

            this.Left = locationX + screenWidth / 2 - this.Width / 2;
            this.Top = SystemParameters.PrimaryScreenHeight / 2 - this.Height / 2;


            OnLoad();

            this.DataContext = userOptions;
            this.Closed += OptionsForm_Closed;

        }



        private void OnLoad()
        {

            userOptions = JsonReader.LoadFromJson<Options>(FilesLocation.JsonOptionsPath);

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using(doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {

                    // Получаем таблицу текстовых стилей
                    TextStyleTable textStyleTable = (TextStyleTable)tr.GetObject(db.TextStyleTableId,OpenMode.ForRead);

                    // Перебираем все стили
                    foreach (ObjectId styleId in textStyleTable)
                    {
                        TextStyleTableRecord styleRecord = (TextStyleTableRecord)tr.GetObject(styleId,OpenMode.ForRead);
                        if(!string.IsNullOrEmpty(styleRecord.Name)) textStyleNames.Add(styleRecord.Name);
                    }

                    // стили мультивыносок
                    DBDictionary dict = tr.GetObject( db.MLeaderStyleDictionaryId,OpenMode.ForRead) as DBDictionary;

                    if (dict != null)
                    {
                        foreach (DBDictionaryEntry entry in dict)
                        {
                            MLeaderStyle style = tr.GetObject(entry.Value, OpenMode.ForRead) as MLeaderStyle;
                            mleaderStyleNames.Add(style.Name);
                        }
                    }
                }
            }

            cb_blockTextStyle.ItemsSource = textStyleNames;
            cb_PKTextStyle.ItemsSource = textStyleNames;
            cb_LineTypeTextStyle.ItemsSource = textStyleNames;
            cb_MleadertStyle.ItemsSource = mleaderStyleNames;

        }



        private void OptionsForm_Closed(object sender, EventArgs e)
        {
            JsonReader.SaveToJson<Options>(userOptions, FilesLocation.JsonOptionsPath);
        }

    }
}
