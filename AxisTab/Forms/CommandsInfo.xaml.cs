using AxisTab;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace AxisTab
{
    /// <summary>
    /// Логика взаимодействия для Window1.xaml

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

            // TODO DataGrid больше не поддерживается. Взамен используйте DataGridView. Подробности см. в https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
            DataGrid grid = sender as DataGrid;
            CommandsInfoData cid = grid.SelectedItem as CommandsInfoData;

            if (cid == null) { return; }

            _optionsForm.Close();
            this.Close();

            DrawingHost.Current.doc?.SendStringToExecute(cid.EngName + " ", true, false, false);

        }


    }
}