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
using ACAD_test;
using System.Xml.Linq;
using System.Windows.Media.Animation;
using System.IO;

namespace TSODD
{
    /// <summary>
    /// Логика взаимодействия для ExportEnd.xaml
    /// </summary>
    public partial class ExportEnd : Window
    {

        Storyboard windowStoryboardStart = new Storyboard();
        Storyboard windowStoryboardEnd = new Storyboard();
        Storyboard textStoryboard = new Storyboard();
        int screenWidth;
        int screenHeight;

        int minValue;
        int maxValue;
        int type = 0;


        private List<string> txtList = new List<string>
        {
            $"Привет, бро! \n Я всё \n сделал",
            "Здравствуй. \n Ведомость\n готова",
            "День добрый.\n Всё \nготово",
            "Здорово! \n Сохранил \n ведомость",
            "Приветствую! \n Всё \n сделано",
            "Салют! \n Ведомость \n выгружена",
            "Hi! Как сам? \n Я всё \n сохранил",
            "Хелло! \n Ведомость \n сохранена"
        };


        public ExportEnd()
        {
            InitializeComponent();

            var hwnd = new WindowInteropHelper(this).Handle;
            var scr = Screen.FromHandle(hwnd);
            screenWidth = scr.Bounds.Width;
            screenHeight = scr.Bounds.Height;

            this.Height = 0;
            this.Width = 0;

            text_textblock.Visibility = Visibility.Collapsed;
            cloud_image.Visibility = Visibility.Collapsed;

            windowStoryboardStart.Completed += WindowStoryboardStart_Completed;
            windowStoryboardEnd.Completed += WindowStoryboardEnd_Completed;

            Random rndm = new Random();
            type = rndm.Next(0, 4);

            StartAnimation();

        }




        private void StartAnimation()
        {
            int height = 0;
            int width = 0;

            switch (type)
            {
                case 0: // снизу 

                    height = 410;
                    width = 400;
                    
                    this.Top = screenHeight;
                    this.Height = 0;
                    this.Width = width;

                    GetImage(cock_image, "cock_1", 20, 180);
                    GetImage(cloud_image, "cloud_left", 10, 10);
                    GetText(25,28);
                    RandomizeFormPosition(true);

                    maxValue = height;
                    minValue = 0;

                    BottomAnimationWindow(windowStoryboardStart, maxValue, minValue, false);

                    break;

                case 1: // снизу 2

                    height = 440;
                    width = 410;

                    this.Top = screenHeight;
                    this.Height = 0;
                    this.Width = width;

                    GetImage(cock_image, "cock_2", 0, 190);
                    GetImage(cloud_image, "cloud_left", 10, 10);
                    GetText(25, 28);
                    RandomizeFormPosition(true);

                    maxValue = height;
                    minValue = 0;

                    BottomAnimationWindow(windowStoryboardStart, maxValue, minValue, false);

                    break;

                case 2: // слева

                    height = 385;
                    width = 385;

                    this.Left = 0;
                    this.Height = height;
                    this.Width = 0;

                    GetImage(cock_image, "cock_4", 0, -300);
                    GetImage(cloud_image, "cloud_right_2", 175, 165);
                    GetText(220, 175);
                    RandomizeFormPosition(false);

                    maxValue = width;
                    minValue = 0;

                    LeftAnimationWindow(windowStoryboardStart, maxValue, minValue, false);

                    break;

                case 3: // справа

                    height = 400;
                    width = 385;

                    this.Left = screenWidth - width;
                    this.Height = height;
                    this.Width = width;

                    GetImage(cock_image, "cock_3", -10, 20);
                    GetImage(cloud_image, "cloud_left_2", 195, 0);
                    GetText(240, 15);
                    RandomizeFormPosition(false);
                    
                    maxValue = width;
                    minValue = 0;

                    RightAnimationWindow(windowStoryboardStart, maxValue, minValue, false);
                    
                    break;

            }
        }




        private void RandomizeFormPosition(bool horizontal)
        {
            Random randomVal = new Random();
            if (horizontal)
            {
                int Xposition = randomVal.Next((int)(screenWidth / 4 - this.Width / 2), (int)(screenWidth * 3 / 4 - this.Width / 2));
                this.Left = Xposition;
            }
            else
            { 
                int Yposition = randomVal.Next((int)(screenHeight / 4 - this.Height / 2), (int)(screenHeight * 3 / 4 - this.Height / 2));
                this.Top = Yposition;
            }
        }


        private void BottomAnimationWindow(Storyboard storyboard, int max, int min, bool invert)
        {
            // Анимация высоты
            DoubleAnimation heightAnimation = new DoubleAnimation
            {
                From = invert? max : min ,
                To = invert ? min : max,
                Duration = TimeSpan.FromSeconds(0.3)
            };

            Storyboard.SetTarget(heightAnimation, this);
            Storyboard.SetTargetProperty(heightAnimation, new PropertyPath(Window.HeightProperty));

            // Анимация позиции
            DoubleAnimation topAnimation = new DoubleAnimation
            {
                From = invert ? screenHeight - max: screenHeight - min,
                To = invert ? screenHeight -  min : screenHeight - max,
                Duration = TimeSpan.FromSeconds(0.3)
            };

            Storyboard.SetTarget(topAnimation, this);
            Storyboard.SetTargetProperty(topAnimation, new PropertyPath(Window.TopProperty));

            // Добавляем в Storyboard
            storyboard.Children.Add(heightAnimation);
            storyboard.Children.Add(topAnimation);
            storyboard.Begin();
        }




        private void LeftAnimationWindow(Storyboard storyboard, int max, int min, bool invert)
        {
            // Анимация высоты
            DoubleAnimation widthAnimation = new DoubleAnimation
            {
                From = invert ? max : min,
                To = invert ? min : max,
                Duration = TimeSpan.FromSeconds(0.3)
            };

            Storyboard.SetTarget(widthAnimation, this);
            Storyboard.SetTargetProperty(widthAnimation, new PropertyPath(Window.WidthProperty));

            // Анимация позиции
            DoubleAnimation leftAnimation = new DoubleAnimation
            {
                From = invert ? 0 : -cock_image.Width,
                To = invert ?  -cock_image.Width : 0,
                Duration = TimeSpan.FromSeconds(0.3)
            };

            Storyboard.SetTarget(leftAnimation, cock_image);
            Storyboard.SetTargetProperty(leftAnimation, new PropertyPath(Window.LeftProperty));

            // Добавляем в Storyboard
            storyboard.Children.Add(widthAnimation);
            storyboard.Children.Add(leftAnimation);
            storyboard.Begin();
        }



        private void RightAnimationWindow(Storyboard storyboard, int max, int min, bool invert)
        {
            // Анимация высоты
            DoubleAnimation widthAnimation = new DoubleAnimation
            {
                From = invert ? max : min,
                To = invert ? min : max,
                Duration = TimeSpan.FromSeconds(0.3)
            };

            Storyboard.SetTarget(widthAnimation, this);
            Storyboard.SetTargetProperty(widthAnimation, new PropertyPath(Window.WidthProperty));

            // Анимация позиции
            DoubleAnimation leftAnimation = new DoubleAnimation
            {
                From = invert ? screenWidth - max : screenWidth - min,
                To = invert ? screenWidth - min : screenWidth - max,
                Duration = TimeSpan.FromSeconds(0.3)
            };

            Storyboard.SetTarget(leftAnimation, this);
            Storyboard.SetTargetProperty(leftAnimation, new PropertyPath(Window.LeftProperty));

            // Добавляем в Storyboard
            storyboard.Children.Add(widthAnimation);
            storyboard.Children.Add(leftAnimation);
            storyboard.Begin();
        }







        // обработчик после первой анимации окна
        private async void WindowStoryboardStart_Completed(object sender, EventArgs e)
        {
            await Task.Delay(300); // пауза 
            text_textblock.Visibility = Visibility.Visible;
            cloud_image.Visibility = Visibility.Visible;

            // теперь запускаем анимацию текста 
            DoubleAnimation doubleAnimationOpacity = new DoubleAnimation()
            {
                From = 0, 
                To = 1,
                Duration = TimeSpan.FromSeconds(0.3)
            };

            Storyboard.SetTarget(doubleAnimationOpacity, cloud_image);
            Storyboard.SetTargetProperty(doubleAnimationOpacity, new PropertyPath(Window.OpacityProperty));
            Storyboard.SetTarget(doubleAnimationOpacity, text_textblock);
            Storyboard.SetTargetProperty(doubleAnimationOpacity, new PropertyPath(Window.OpacityProperty));

            // Добавляем в Storyboard
            textStoryboard.Children.Add(doubleAnimationOpacity);
            //await Task.Run(() => textStoryboard.Begin()); 
            textStoryboard.Begin();
            textStoryboard.Children.Clear();

            await Task.Delay(1500);     //  пауза

            doubleAnimationOpacity = new DoubleAnimation()
            {
                From = 1,
                To = 0,
                Duration = TimeSpan.FromSeconds(0.3)
            };

            Storyboard.SetTarget(doubleAnimationOpacity, cloud_image);
            Storyboard.SetTargetProperty(doubleAnimationOpacity, new PropertyPath(Window.OpacityProperty));
            Storyboard.SetTarget(doubleAnimationOpacity, text_textblock);
            Storyboard.SetTargetProperty(doubleAnimationOpacity, new PropertyPath(Window.OpacityProperty));

            // Добавляем в Storyboard
            textStoryboard.Children.Add(doubleAnimationOpacity);
            //textStoryboard.Begin();

            await Task.Delay(300); // пауза 
            text_textblock.Visibility = Visibility.Collapsed;
            cloud_image.Visibility = Visibility.Collapsed;

            EndAnimation();

        }


        private async void EndAnimation()
        {
            switch (type)
            {
                case 0: case 1: BottomAnimationWindow(windowStoryboardEnd, maxValue, minValue, true); break;
                case 2: LeftAnimationWindow(windowStoryboardEnd, maxValue, minValue, true); break;
                case 3: RightAnimationWindow(windowStoryboardEnd, maxValue, minValue, true); break;
            }
        }


        private void WindowStoryboardEnd_Completed(object sender, EventArgs e)
        {
            this.Close();
        }



        // настройка картинки
        void GetImage(Image element, string name, int top, int left)
        {
            BitmapImage image = RibbonInitializer.Instance.LoadImage($"pack://application:,,,/TSODD;component/images/{name}.png");
            image.Freeze();
            element.Source = image;
            element.Width = image.PixelWidth;
            element.Height = image.PixelHeight;
            Canvas.SetLeft(element, left);
            Canvas.SetTop(element, top);
        }

        // настройка положения текста
        void GetText(int top, int left)
        {
            Random random = new Random();
            int number = random.Next(0, txtList.Count - 1);
            text_textblock.Text = txtList[number];
            Canvas.SetLeft(text_textblock, left);
            Canvas.SetTop(text_textblock, top);
        }








    }
}
