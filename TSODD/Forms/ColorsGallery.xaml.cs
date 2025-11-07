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

namespace TSODD.forms
{

    public class ColorItem
    {
        public string Text { get; set; }
        public System.Windows.Media.Color Color_1 { get; set; }
        public System.Windows.Media.Color Color_2 { get; set; }
        public short Index { get; set; }
    }



    /// <summary>
    /// Логика взаимодействия для GroupsAddForm.xaml
    /// </summary>
    public partial class ColorsGallery : Window
    {
        private ObservableCollection<ColorItem> colors = new ObservableCollection<ColorItem>();
        private LineTypeForm _lineTypeForm;

        public ColorsGallery(LineTypeForm lineTypeForm)
        {
            InitializeComponent();
            _lineTypeForm = lineTypeForm;
            this.Top = lineTypeForm.Top;
            this.Left = lineTypeForm.Left + lineTypeForm.Width;

            lv_Colors.ItemsSource = colors;
            FillColorsGallery();
        }

        private void FillColorsGallery()
        {
            for (short i = 0; i < 256; i++)
            {
                var colorAcad = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, i);
                var colorWpf = System.Windows.Media.Color.FromArgb(colorAcad.ColorValue.A, colorAcad.ColorValue.R, colorAcad.ColorValue.G, colorAcad.ColorValue.B);
                string txt = $"цвет {i}";
                colors.Add(new ColorItem { Text = txt, Color_1 = colorWpf, Color_2 = colorWpf, Index = i });
            }
        }

        private void lv_Colors_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var itemColor = lv_Colors.SelectedItem as ColorItem;
            var colorAcad = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, itemColor.Index);
            string txt = $"цвет {itemColor.Index} ARGB = ({colorAcad.ColorValue.A},{colorAcad.ColorValue.R},{colorAcad.ColorValue.G},{colorAcad.ColorValue.B})";
            tb_Text.Text = txt;
        }

        private void ButtonAdd_Click(object sender, RoutedEventArgs e)
        {
            var itemColor = lv_Colors.SelectedItem as ColorItem;
            if (itemColor == null) return;

            _lineTypeForm.selectedColorItem = itemColor;
            _lineTypeForm.brd_ColorBorder_1.Background = new SolidColorBrush(itemColor.Color_1);
            _lineTypeForm.brd_ColorBorder_2.Background = new SolidColorBrush(itemColor.Color_2);
         
            this.Close();
        }

    
        private void ButtonExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }


    }
}
