using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Input;
using AxisTAb;
using System.Windows.Threading;
using System.Windows;
using System.Diagnostics;
using System.Runtime.InteropServices;
using AxisTab;


namespace AxisTAb
{
    public partial class RibbonInitializer : IExtensionApplication
    {
        public static RibbonInitializer Instance { get; private set; }
        private RibbonCombo axisCombo = new RibbonCombo();

        private static readonly HashSet<IntPtr> _dbIntPtr = new HashSet<IntPtr>();
        private readonly HashSet<ObjectId> _dontDeleteMe = new HashSet<ObjectId>();
        private readonly HashSet<ObjectId> _deleteMe = new HashSet<ObjectId>();
        public DispatcherTimer timer;

        public bool inactivity = JsonReader.LoadFromJson<Options>(FilesLocation.JsonOptionsPath).Inactivity;
        public double currentTimeSpan = JsonReader.LoadFromJson<Options>(FilesLocation.JsonOptionsPath).InactivityTimeSpan;
        
        AnimationForm animationForm;

        public void Initialize()
        {
            //return;

            Instance = this; // для того, что бы можно было методы вызывать
            ComponentManager.ItemInitialized += OnItemInitialized;

            // подписываемся на документ
            var dm = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            dm.DocumentActivated += Dm_DocumentActivated;
            dm.DocumentToBeDestroyed += Dm_DocumentToBeDestroyed;


            // таймер для отслеживания бездействия пользователя
            timer= new DispatcherTimer
            {
                Interval = TimeSpan.FromMinutes(currentTimeSpan)   
            };

            timer.Tick += _timer_Tick;

        }


        // перезапускает таймер отслеживания
        private void RestartTimer()
        {
            timer.Stop();
            timer.Start();
            thisWindowActive = true;

            if (animationForm != null)
            {
                animationForm.RunClosingAnimation();
                animationForm = null;
            }
        }


        // таймер сработал пора показывать анимацию

        private bool thisWindowActive;
        private void _timer_Tick(object sender, EventArgs e)
        {

            if(!inactivity) return;

            if (!WindowAcadActivity.IsAutoCADActive())
            {
                RestartTimer();
                thisWindowActive = false;
                return;
            }

            if (animationForm == null && thisWindowActive)
            {
                animationForm = new AnimationForm();
                animationForm.Show();
            }

        }

        public void Terminate()
        {
            Instance = null;
            ComponentManager.ItemInitialized -= OnItemInitialized;

            // отписываемся от документа
            var dm = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            dm.DocumentActivated -= Dm_DocumentActivated;
            dm.DocumentToBeDestroyed -= Dm_DocumentToBeDestroyed;

            // отпишемся от всего 
            foreach (Autodesk.AutoCAD.ApplicationServices.Document d in dm)
            {
                var db = d.Database;
                var key = db.UnmanagedObject;

                db.ObjectErased -= Db_ObjectErased;
                d.CommandEnded -= MdiActiveDocument_CommandEnded;
                d.CommandCancelled -= MdiActiveDocument_CommandEnded;
                d.Editor.LeavingQuiescentState -= Editor_LeavingQuiescentState;
                d.ViewChanged -= Document_ViewChanged;

                _dbIntPtr.Remove(key);
            }
        }

        private void OnItemInitialized(object sender, EventArgs e)
        {
            if (ComponentManager.Ribbon != null)
            {
                ComponentManager.ItemInitialized -= OnItemInitialized;
                axisCombo.CurrentChanged += AxisCombo_CurrentChanged;

                AddRibbonPanel();

            }
        }


        private void AddRibbonPanel()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            RibbonControl ribbon = ComponentManager.Ribbon;
            if (ribbon == null)
                return;

            // Проверяем, есть ли уже вкладка
            RibbonTab existingTab = ribbon.Tabs.FirstOrDefault(t => t.Id == "ACAD_AXISTAB");
            if (existingTab != null)
                return;

            // Вкладка
            RibbonTab tab = new RibbonTab
            {
                Title = "Оси",
                Id = "ACAD_AXISTAB"
            };
            ribbon.Tabs.Add(tab);


            /* ************************************************         ОСИ            ************************************************ */

            RibbonPanelSource panelSourceAxis = new RibbonPanelSource
            {
                Title = "Ось"
            };

            RibbonPanel panel_1 = new RibbonPanel
            {
                Source = panelSourceAxis
            };
            tab.Panels.Add(panel_1);

            // Большая кнопка NEWAXIS
            var bt_newAxis = new RibbonButton
            {
                Name = "NEWAXIS",
                Text = "Новая ось",
                ShowText = true,
                LargeImage = LoadImage("pack://application:,,,/AxisTAb;component/images/axis_new.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,

                CommandHandler = new RelayCommandHandler(() =>
                {
                    DrawingHost.Current.doc?.SendStringToExecute("IA_NEW_AXIS ", true, false, false);
                })
            };
            panelSourceAxis.Items.Add(bt_newAxis);
            panelSourceAxis.Items.Add(new RibbonSeparator());

            var rowAxis = new RibbonRowPanel();

            // список combobox с осями 
            axisCombo.Text = " Текущая:";
            axisCombo.ShowText = true;
            axisCombo.Width = 150;
            rowAxis.Items.Add(axisCombo);

            // Перенос на новую строку внутри этой же группы
            rowAxis.Items.Add(new RibbonRowBreak());

            //  Кнопка изменения имени оси
            var bt_changeName = new RibbonButton
            {
                Text = "Имя оси",
                ShowText = true,
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
                Image = LoadImage("pack://application:,,,/AxisTAb;component/images/axis_name.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    DrawingHost.Current.doc?.SendStringToExecute("IA_SET_AXIS_NAME ", true, false, false);
                })
            };

            // Кнопка изменения начальной точки
            var bt_startPoint = new RibbonButton
            {
                Text = "Начальная точка оси",
                ShowText = true,
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
                Image = LoadImage("pack://application:,,,/AxisTAb;component/images/axis_startPoint.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    DrawingHost.Current.doc?.SendStringToExecute("IA_SET_AXIS_START_POINT ", true, false, false);
                })
            };


            // Кнопки рядом + разделитель между ними
            rowAxis.Items.Add(bt_changeName);

            rowAxis.Items.Add(new RibbonRowBreak());

            rowAxis.Items.Add(bt_startPoint);

            // добавляем группу "ось" в панель
            panelSourceAxis.Items.Add(rowAxis);

            panelSourceAxis.Items.Add(new RibbonSeparator());

            // Кнопка Отбить ПК
            var bt_setPK = new RibbonButton
            {
                Text = "Назначить ПК",
                ShowText = true,
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
                LargeImage = LoadImage("pack://application:,,,/AxisTAb;component/images/axis_setPK.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    DrawingHost.Current.doc?.SendStringToExecute("IA_SET_PK ", true, false, false); 
                })
            };

            // Кнопка Отбить ПК
            var bt_getPK = new RibbonButton
            {
                Text = "Получить ПК",
                ShowText = true,
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
                LargeImage = LoadImage("pack://application:,,,/AxisTAb;component/images/axis_getPK.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    DrawingHost.Current.doc?.SendStringToExecute("IA_GET_PK ", true, false, false);
                })
            };

            panelSourceAxis.Items.Add(bt_setPK);
            panelSourceAxis.Items.Add(bt_getPK);



            /* ************************************************          НАСТРОЙКИ и БД            ************************************************ */

            RibbonPanelSource panelSourceOptions = new RibbonPanelSource
            {
                Title = "Настройки"
            };

            RibbonPanel panel_5 = new RibbonPanel
            {
                Source = panelSourceOptions
            };
            tab.Panels.Add(panel_5);

            var bt_options = new RibbonButton
            {
                Name = "",
                Text = "Настройки",
                ShowText = true,
                LargeImage = LoadImage("pack://application:,,,/AxisTAb;component/images/options.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,
                CommandHandler = new RelayCommandHandler(() =>
                {
                    DrawingHost.Current.doc?.SendStringToExecute("IA_OPTIONS_AXIS ", true, false, false);
                })
            };
           
            panelSourceOptions.Items.Add(bt_options);
            panelSourceOptions.Items.Add(new RibbonSeparator());
        }
    }






    public class RelayCommandHandler : ICommand
    {
        private readonly Action _action;

        public RelayCommandHandler(Action action)
        {
            _action = action;
        }

        public event EventHandler CanExecuteChanged
        {
            add { }
            remove { }
        }

        public bool CanExecute(object parameter) => true;

        public void Execute(object parameter)
        {
            _action?.Invoke();
        }
    }



























}
