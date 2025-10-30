using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Interop;

namespace TSODD
{
    /// <summary>
    /// Логика взаимодействия для wwindow.xaml
    /// </summary>
    public partial class wwindow : Window
    {
        public wwindow()
        {
            InitializeComponent();

            var hwnd = new WindowInteropHelper(this).Handle;
            var scr = Screen.FromHandle(hwnd);
            int w = scr.Bounds.Width;
            int h = scr.Bounds.Height;

            this.Left = w-10;

        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {

            for (int i = 0; i < 30; i++)
            {
                await Task.Delay(8);
                this.Left -=8;



            }

        }




    }
}
