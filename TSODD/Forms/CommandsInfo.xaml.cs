using System;
using System.Collections.Generic;
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

namespace TSODD.Forms
{

    public class CommandsInfoData 
    {
        public string EngName { get; set; }
        public string RusName { get; set; }
        public string Description { get; set; }
        public string ImageSource { get; set; }
        public int ImageSize { get; set; }
    }




    /// <summary>
    /// Логика взаимодействия для CommandsInfo.xaml
    /// </summary>
    public partial class CommandsInfo : Window
    {
        OptionsForm _optionsForm;

        public CommandsInfo(OptionsForm optionsForm)
        {
            InitializeComponent();
            _optionsForm = optionsForm;
            this.Top = optionsForm.Top;
            this.Left = optionsForm.Left + 370;

            this.Closed += CommandsInfo_Closed;
            this.LocationChanged += CommandsInfo_LocationChanged;
        }

        private void CommandsInfo_LocationChanged(object sender, EventArgs e)
        {
            // перетаскиваем основное окно настроек
            if (_optionsForm.dragFlag) return;
            _optionsForm.dragFlag = true; 

            _optionsForm.Top = this.Top;
            _optionsForm.Left = this.Left - _optionsForm.Width;

            _optionsForm.dragFlag = false;
        }



        private void CommandsInfo_Closed(object sender, EventArgs e)
        {
           _optionsForm.commandsInfo = null;
        }

        private void grid_commands_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {

            DataGrid grid = sender as DataGrid;
            CommandsInfoData cid = grid.SelectedItem as CommandsInfoData;

            if (cid == null) { return; }

            _optionsForm.Close();
            this.Close();

            TsoddHost.Current.doc?.SendStringToExecute(cid.EngName+" ", true, false, false);

        }


    }
}
