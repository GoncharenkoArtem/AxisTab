using Autodesk.AutoCAD.Interop.Common;
using AxisTAb;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace AxisTab
{
    /// <summary>
    /// Логика взаимодействия для AnimationForm.xaml
    /// </summary>
    public partial class AnimationForm : Window
    {
        Storyboard windowStoryboardStart = new Storyboard();
        Storyboard windowStoryboardEnd = new Storyboard();
        Storyboard textStoryboardEnd = new Storyboard();


        double locationX;
        double locationY;
        double screenWidth;
        double screenHeight;

        double minValue;
        double maxValue;
        int type = 0;

        DispatcherTimer _timer;
        private int startTime  =(int)(JsonReader.LoadFromJson<Options>(FilesLocation.JsonOptionsPath).InactivityTimeSpan*60);
        private int currentTime;

        private List<string> txtList = new List<string>
        {
            $"Спишь?\nПросыпайся",
            "Пришло\nвремя клацать",
            "Давай\nклац-клац",
            "Всё хорошо?\n Поработаем?",
            "Давай работать, а?",
            "Работать \n будем?",
            "Ну,так и\nбудем сидеть?",
            "Жми, давай\n скорее!",
            "Не работает \n и всё тут!",
            "Так мы денег\nне заработаем",
        };


        public AnimationForm()
        {
            InitializeComponent();

            var app = Autodesk.AutoCAD.ApplicationServices.Application.MainWindow;

            var doc = DrawingHost.Current.doc;

            locationX = app.DeviceIndependentLocation.X;
            locationY = app.DeviceIndependentLocation.Y;
            screenWidth = app.DeviceIndependentSize.Width;
            screenHeight = app.DeviceIndependentSize.Height;

            this.Height = 0;
            this.Width = 0;

            text_textblock.Visibility = Visibility.Collapsed;
            text_timeblock.Visibility = Visibility.Collapsed;
            text_secblock.Visibility = Visibility.Collapsed;
            cloud_image.Visibility = Visibility.Collapsed;

            windowStoryboardStart.Completed += WindowStoryboardStart_Completed;
            textStoryboardEnd.Completed += TextStoryboardEnd_Completed;
            windowStoryboardEnd.Completed += WindowStoryboardEnd_Completed; 

            this.MouseLeftButtonDown += AnimationForm_MouseLeftButtonDown;

            Random rndm = new Random();
            type = rndm.Next(0, 4);

            // таймер для обновления счетчика
            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromSeconds(1);
            _timer.Stop();
            _timer.Tick += _timer_Tick;


            double currentFullTime = JsonReader.LoadFromJson<Options>(FilesLocation.JsonOptionsPath).InactivityFullTime;
            currentTime = (int)(currentFullTime*60 - startTime-5);

            // старт анимации
            StartAnimation();
        }



        // счетчик сработал
        private void _timer_Tick(object sender, EventArgs e)
        {
            if (currentTime > 0)
            {
                currentTime -= 1;
                text_timeblock.Text = $"{currentTime}";
            }
            else
            {
                text_textblock.Text = "ПОТРАЧЕНО";

                Storyboard fuckUpstory = new Storyboard();

                // Анимация позиции
                DoubleAnimation topAnimation = new DoubleAnimation
                {
                    From = Canvas.GetTop(text_textblock),
                    To = Canvas.GetTop(text_textblock)+40,
                    Duration = TimeSpan.FromSeconds(0.3)
                };

                Storyboard.SetTarget(topAnimation, text_textblock);
                Storyboard.SetTargetProperty(topAnimation, new PropertyPath(Window.TopProperty));
                fuckUpstory.Children.Add(topAnimation);
                fuckUpstory.Begin();


                text_timeblock.Visibility = Visibility.Collapsed;
                text_secblock.Visibility = Visibility.Collapsed;
                _timer.Stop();
            }
            
            if (currentTime < 26 && currentTime > 16) { text_timeblock.Foreground = Brushes.Orange; }
            if (currentTime < 16 && currentTime > 6) { text_timeblock.Foreground = Brushes.Salmon; }
            if (currentTime < 6) { text_timeblock.Foreground = Brushes.Red; }

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

                    this.Top = locationY + screenHeight;
                    this.Height = 0;
                    this.Width = width;

                    GetImage(cock_image, "cock_1", 60, 180);
                    GetImage(cloud_image, "cloud_left", 50, 10);
                    GetText(60, 28);
                    RandomizeFormPosition(true);

                    maxValue = height;
                    minValue = 0;

                    BottomAnimationWindow(windowStoryboardStart, maxValue, minValue, false);

                    break;

                case 1: // снизу 2

                    height = 440;
                    width = 410;

                    this.Top = locationY + screenHeight;
                    this.Height = 0;
                    this.Width = width;

                    GetImage(cock_image, "cock_2", 40, 190);
                    GetImage(cloud_image, "cloud_left", 50, 10);
                    GetText(60, 28);
                    RandomizeFormPosition(true);

                    maxValue = height;
                    minValue = 0;

                    BottomAnimationWindow(windowStoryboardStart, maxValue, minValue, false);

                    break;

                case 2: // слева

                    height = 385;
                    width = 385;

                    this.Left = locationX;
                    this.Height = height;
                    this.Width = 0;

                    GetImage(cock_image, "cock_4", 0, -300);
                    GetImage(cloud_image, "cloud_right_2", 175, 165);
                    GetText(220, 172);
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

                    GetImage(cock_image, "cock_3", -10, 35);
                    GetImage(cloud_image, "cloud_left_2", 195, 15);
                    GetText(240, 30);
                    RandomizeFormPosition(false);

                    maxValue = width;
                    minValue = 0;

                    RightAnimationWindow(windowStoryboardStart, maxValue, minValue, false);

                    break;
            }
            _timer.Start();
        }


        private void RandomizeFormPosition(bool horizontal)
        {
            Random randomVal = new Random();
            if (horizontal)
            {
                int Xposition = randomVal.Next((int)((locationX + screenWidth / 4) - this.Width / 2), (int)((locationX + screenWidth * 3 / 4) - this.Width / 2));
                this.Left = Xposition;
            }
            else
            {
                int Yposition = randomVal.Next((int)((locationY + screenHeight / 4) - this.Height / 2), (int)((locationY + screenHeight * 3 / 4) - this.Height / 2));
                this.Top = Yposition;
            }
        }


        private void BottomAnimationWindow(Storyboard storyboard, double max, double min, bool invert)
        {
            // Анимация высоты
            DoubleAnimation heightAnimation = new DoubleAnimation
            {
                From = invert ? max : min,
                To = invert ? min : max,
                Duration = TimeSpan.FromSeconds(0.3)
            };

            Storyboard.SetTarget(heightAnimation, this);
            Storyboard.SetTargetProperty(heightAnimation, new PropertyPath(Window.HeightProperty));

            // Анимация позиции
            DoubleAnimation topAnimation = new DoubleAnimation
            {
                From = invert ? screenHeight + locationY - max : screenHeight + locationY - min,
                To = invert ? screenHeight + locationY - min : screenHeight + locationY - max,
                Duration = TimeSpan.FromSeconds(0.3)
            };

            Storyboard.SetTarget(topAnimation, this);
            Storyboard.SetTargetProperty(topAnimation, new PropertyPath(Window.TopProperty));

            // Добавляем в Storyboard
            storyboard.Children.Add(heightAnimation);
            storyboard.Children.Add(topAnimation);
            storyboard.Begin();

        }


        private void LeftAnimationWindow(Storyboard storyboard, double max, double min, bool invert)
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
                To = invert ? -cock_image.Width : 0,
                Duration = TimeSpan.FromSeconds(0.3)
            };

            Storyboard.SetTarget(leftAnimation, cock_image);
            Storyboard.SetTargetProperty(leftAnimation, new PropertyPath(Window.LeftProperty));

            // Добавляем в Storyboard
            storyboard.Children.Add(widthAnimation);
            storyboard.Children.Add(leftAnimation);
            storyboard.Begin();
        }


        private void RightAnimationWindow(Storyboard storyboard, double max, double min, bool invert)
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
                From = invert ? locationX + screenWidth - max : locationX + screenWidth - min,
                To = invert ? locationX + screenWidth - min : locationX + screenWidth - max,
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
            text_timeblock.Visibility = Visibility.Visible;
            text_secblock.Visibility = Visibility.Visible;

            cloud_image.Visibility = Visibility.Visible;
            text_timeblock.Text = $"{currentTime}";


            Storyboard textStoryboard =new Storyboard();

            var doubleAnimationOpacity_1 = new DoubleAnimation()
            {
                From = 1,
                To = 0,
                Duration = TimeSpan.FromSeconds(0.3)
            };

            var doubleAnimationOpacity_2 = new DoubleAnimation()
            {
                From = 1,
                To = 0,
                Duration = TimeSpan.FromSeconds(0.3)
            };

            Storyboard.SetTarget(doubleAnimationOpacity_1, cloud_image);
            Storyboard.SetTargetProperty(doubleAnimationOpacity_1, new PropertyPath(Window.OpacityProperty));
            textStoryboardEnd.Children.Add(doubleAnimationOpacity_1);

            Storyboard.SetTarget(doubleAnimationOpacity_2, text_textblock);
            Storyboard.SetTargetProperty(doubleAnimationOpacity_2, new PropertyPath(Window.OpacityProperty));
            textStoryboardEnd.Children.Add(doubleAnimationOpacity_2);

            textStoryboard.Begin();
        }


        private void TextStoryboardEnd_Completed(object sender, EventArgs e)
        {
            EndAnimation();
        }


        private void EndAnimation()
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


        private void AnimationForm_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            RunClosingAnimation();
        }


        // анимация завершения 
        public async void RunClosingAnimation()
        {

            if (currentTime > 0)
            {
                text_timeblock.Visibility = Visibility.Collapsed;
                text_secblock.Visibility = Visibility.Collapsed;
                text_textblock.Text = "УСПЕЛИ";
                Canvas.SetTop(text_textblock, Canvas.GetTop(text_textblock) + 40);
                await Task.Delay(300); // пауза  
            }

            var doubleAnimationOpacity_1 = new DoubleAnimation()
            {
                From = 1,
                To = 0,
                Duration = TimeSpan.FromSeconds(0.3)
            };

            var doubleAnimationOpacity_2 = new DoubleAnimation()
            {
                From = 1,
                To = 0,
                Duration = TimeSpan.FromSeconds(0.3)
            };

            Storyboard.SetTarget(doubleAnimationOpacity_1, cloud_image);
            Storyboard.SetTargetProperty(doubleAnimationOpacity_1, new PropertyPath(Window.OpacityProperty));
            textStoryboardEnd.Children.Add(doubleAnimationOpacity_1);

            Storyboard.SetTarget(doubleAnimationOpacity_2, text_textblock);
            Storyboard.SetTargetProperty(doubleAnimationOpacity_2, new PropertyPath(Window.OpacityProperty));
            textStoryboardEnd.Children.Add(doubleAnimationOpacity_2);

            textStoryboardEnd.Begin();

        }







        // настройка картинки
        void GetImage(Image element, string name, int top, int left)
        {
            BitmapImage image = RibbonInitializer.Instance.LoadImage($"pack://application:,,,/AxisTAb;component/images/{name}.png");
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
            int number = random.Next(0, txtList.Count);
            text_textblock.Text = txtList[number];
            
            // текст подписи
            Canvas.SetLeft(text_textblock, left);
            Canvas.SetTop(text_textblock, top);

            // текст времени
            Canvas.SetLeft(text_timeblock, left+32);
            Canvas.SetTop(text_timeblock, top+72);

            //текст подписи "сек"
            Canvas.SetLeft(text_secblock, left + 85);
            Canvas.SetTop(text_secblock, top + 72);
        }

    }
}