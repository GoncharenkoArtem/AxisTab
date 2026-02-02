using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Input;
using TSODD;
using TSODD.forms;
using TSODD.Forms;


namespace TSODD
{
    public partial class RibbonInitializer : IExtensionApplication
    {
        public static RibbonInitializer Instance { get; private set; }
        private RibbonCombo axisCombo = new RibbonCombo();

        public RibbonSplitButton splitStands = new RibbonSplitButton();
        public RibbonSplitButton splitMarksLineTypes = new RibbonSplitButton();
        public RibbonButton quickProperties = new RibbonButton();


        public RibbonRowPanel rowLineType = new RibbonRowPanel();
        private RibbonCombo comboLineTypePattern_1 = new RibbonCombo();
        private RibbonCombo comboLineTypePattern_2 = new RibbonCombo();
        private RibbonCombo comboLineTypeWidth_1 = new RibbonCombo();
        private RibbonCombo comboLineTypeWidth_2 = new RibbonCombo();
        private RibbonCombo comboLineTypeColor_1 = new RibbonCombo();
        private RibbonCombo comboLineTypeColor_2 = new RibbonCombo();
        private RibbonCombo comboLineTypeOffset = new RibbonCombo();
        private RibbonLabel labelLineType = new RibbonLabel();

        public RibbonPanelSource panelSourceMarks;

        private static readonly HashSet<IntPtr> _dbIntPtr = new HashSet<IntPtr>();
        private readonly HashSet<ObjectId> _dontDeleteMe = new HashSet<ObjectId>();
        private readonly HashSet<ObjectId> _deleteMe = new HashSet<ObjectId>();

        private bool _marksLineTypeFlag = true;
        public bool readyToDeleteEntity = true;
        public bool quickPropertiesOn = false;
        public SelectionFormBlocks selectioForm = null;

        public void Initialize()
        {
            //return;

            Instance = this; // для того, что бы можно было методы вызывать
            ComponentManager.ItemInitialized += OnItemInitialized;

            // подписываемся на документ
            var dm = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            dm.DocumentActivated += Dm_DocumentActivated;
            dm.DocumentToBeDestroyed += Dm_DocumentToBeDestroyed;

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
                d.Editor.PromptForSelectionEnding -= Editor_PromptForSelectionEnding;

                _dbIntPtr.Remove(key);
            }
        }

        private void OnItemInitialized(object sender, EventArgs e)
        {
            if (ComponentManager.Ribbon != null)
            {
                ComponentManager.ItemInitialized -= OnItemInitialized;
                axisCombo.CurrentChanged += AxisCombo_CurrentChanged;
                splitStands.CurrentChanged += SplitStands_CurrentChanged;

                AddRibbonPanel();

                // предвыбор значений
                //TsoddBlock.PreSelectOfGroups();

                // внутренний метод, который первоначально настраивает RibbonSplitButton
                void InitializeSplitButtons(RibbonSplitButton split, string name)
                {
                    split.Text = name;
                    split.Size = RibbonItemSize.Large;
                    split.Orientation = Orientation.Vertical;
                    split.IsSplit = true;
                    split.ShowText = true;
                    split.Width = 80;
                }

                // настраиваем контрол со стойками
                InitializeSplitButtons(splitStands, "Стойки");
                
                // заполняем 
                //FillBlocksMenu(splitStands, "STAND");

                // обновляем  элементы LineType
                ListOfMarksLinesLoad(200, 20);
                LineTypeReader.RefreshLineTypesInAcad();

            }
        }


        private void AddRibbonPanel()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            RibbonControl ribbon = ComponentManager.Ribbon;
            if (ribbon == null)
                return;

            // Проверяем, есть ли уже вкладка
            RibbonTab existingTab = ribbon.Tabs.FirstOrDefault(t => t.Id == "ACAD_TSODD");
            if (existingTab != null)
                return;

            // Вкладка
            RibbonTab tab = new RibbonTab
            {
                Title = "ТСОДД",
                Id = "ACAD_TSODD"
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
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/axis_new.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,

                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("NEW_AXIS ", true, false, false);
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
                Image = LoadImage("pack://application:,,,/TSODD;component/images/axis_name.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("SET_AXIS_NAME ", true, false, false);
                })
            };

            // Кнопка изменения начальной точки
            var bt_startPoint = new RibbonButton
            {
                Text = "Начальная точка оси",
                ShowText = true,
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
                Image = LoadImage("pack://application:,,,/TSODD;component/images/axis_startPoint.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("SET_AXIS_START_POINT ", true, false, false);
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
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/axis_setPK.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("SET_PK ", true, false, false); 
                })
            };

            // Кнопка Отбить ПК
            var bt_getPK = new RibbonButton
            {
                Text = "Получить ПК",
                ShowText = true,
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/axis_getPK.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("GET_PK ", true, false, false);
                })
            };

            panelSourceAxis.Items.Add(bt_setPK);
            panelSourceAxis.Items.Add(bt_getPK);

            /* ************************************************           СТОЙКИ            ************************************************ */

            //RibbonPanelSource panelSourceStands = new RibbonPanelSource
            //{
            //    Title = "Стойки"
            //};

            //RibbonPanel panel_2 = new RibbonPanel
            //{
            //    Source = panelSourceStands
            //};
            //tab.Panels.Add(panel_2);

            //// добавляем на ribbon
            //panelSourceStands.Items.Add(splitStands);

            //// кнопка перепривязки стоек к оси
            //RibbonButton axisBinding = new RibbonButton
            //{
            //    Text = "Привязать к оси",
            //    ShowText = true,
            //    Size = RibbonItemSize.Large,
            //    Orientation = Orientation.Vertical,
            //    LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/stand_bindToAxis.png"),
            //    CommandHandler = new RelayCommandHandler(() =>
            //    {
            //        TsoddHost.Current.doc?.SendStringToExecute("BIND_TO_AXIS ", true, false, false);

            //    })
            //};

            //panelSourceStands.Items.Add(axisBinding);


            /* ************************************************           БЛОКИ            ************************************************ */

            RibbonPanelSource panelSourceBlocks = new RibbonPanelSource
            {
                Title = "Блоки"
            };

            RibbonPanel panel_3 = new RibbonPanel
            {
                Source = panelSourceBlocks
            };
            tab.Panels.Add(panel_3);

            // вставка блока
            RibbonButton insertBlock = new RibbonButton
            {
                Text = "Вставить блок",
                ShowText = true,
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/block_insert.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("INSERT_TSODD_BLOCK ", true, false, false);
                })
            };
            panelSourceBlocks.Items.Add(insertBlock);

            // вставка блока
            RibbonButton userBlock = new RibbonButton
            {
                Text = "Пользовательский",
                ShowText = true,
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/block_user.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("USER_MARK_BLOCK ", true, false, false);
                })
            };
            panelSourceBlocks.Items.Add(userBlock);

   
            /* ************************************************          РАЗМЕТКА            ************************************************ */

            panelSourceMarks = new RibbonPanelSource
            {
                Title = "Линейная разметка"
            };

            RibbonPanel panel_4 = new RibbonPanel
            {
                Source = panelSourceMarks
            };
            tab.Panels.Add(panel_4);

            // типы линий

            splitMarksLineTypes.Width = 60;
            splitMarksLineTypes.Size = RibbonItemSize.Large;
            splitMarksLineTypes.IsSplit = true;
            panelSourceMarks.Items.Add(splitMarksLineTypes);

            // первый тип линии
            var rowLineType_1 = new RibbonRowPanel();
            comboLineTypePattern_1.Width = 105;
            comboLineTypeWidth_1.MinWidth = 50;
            comboLineTypeWidth_1.Width = 50;
            comboLineTypeWidth_1.Text = "";
            comboLineTypeWidth_1.ShowText = true;
            comboLineTypeColor_1.MinWidth = 50;
            comboLineTypeColor_1.Width = 50;
            comboLineTypeColor_1.Text = "";
            comboLineTypeColor_1.ShowText = true;
            rowLineType_1.Items.Add(comboLineTypePattern_1);
            rowLineType_1.Items.Add(comboLineTypeWidth_1);
            rowLineType_1.Items.Add(comboLineTypeColor_1);
            rowLineType.Items.Add(rowLineType_1);
            rowLineType.Items.Add(new RibbonRowBreak());
            // второй тип линии
            var rowLineType_2 = new RibbonRowPanel();
            comboLineTypePattern_2.Width = 105;
            comboLineTypeWidth_2.MinWidth = 50;
            comboLineTypeWidth_2.Width = 50;
            comboLineTypeWidth_2.Text = "";
            comboLineTypeWidth_2.ShowText = true;
            comboLineTypeColor_2.MinWidth = 50;
            comboLineTypeColor_2.Width = 50;
            comboLineTypeColor_2.Text = "";
            comboLineTypeColor_2.ShowText = true;
            rowLineType_2.Items.Add(comboLineTypePattern_2);
            rowLineType_2.Items.Add(comboLineTypeWidth_2);
            rowLineType_2.Items.Add(comboLineTypeColor_2);
            rowLineType.Items.Add(rowLineType_2);
            rowLineType.Items.Add(new RibbonRowBreak());
            // расстояние между линиями
            var rowLineType_3 = new RibbonRowPanel();
            labelLineType.Text = "расстояние между линиями";
            labelLineType.Width = 159;
            rowLineType_3.Items.Add(labelLineType);
            comboLineTypeOffset.Width = 46;
            comboLineTypeOffset.MinWidth = 46;
            rowLineType_3.Items.Add(comboLineTypeOffset);
            rowLineType.Items.Add(rowLineType_3);

            panelSourceMarks.Items.Add(rowLineType);

            panelSourceMarks.Items.Add(new RibbonSeparator());

            var bt_invertLineType = new RibbonButton
            {
                Name = "",
                Text = "Инвертировать",
                ShowText = true,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/lineType_invertLines.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("MARK_LINE_INVERT ", true, false, false);
                })
            };
            panelSourceMarks.Items.Add(bt_invertLineType);


            var bt_invertTextPosition = new RibbonButton
            {
                Name = "",
                Text = "Переставить текст",
                ShowText = true,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/lineType_invertText.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("MARK_LINE_TEXT_INVERT ", true, false, false);
                })
            };
            panelSourceMarks.Items.Add(bt_invertTextPosition);


            /* ************************************************          НАСТРОЙКИ и БД            ************************************************ */

            RibbonPanelSource panelSourceOptions = new RibbonPanelSource
            {
                Title = "Настройки и БД"
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
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/options.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("OPTIONS_TSODD ", true, false, false);
                })
            };
            panelSourceOptions.Items.Add(bt_options);

            var bt_addBlockToBD = new RibbonButton
            {
                Name = "",
                Text = "Загрузить блок",
                ShowText = true,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/blocks_loadToBD.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("LOAD_BLOCK_TO_DB ", true, false, false);
                })
            };
            panelSourceOptions.Items.Add(bt_addBlockToBD);

            var bt_addLineTypeToBD = new RibbonButton
            {
                Name = "",
                Text = "Создать разметку",
                ShowText = true,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/lineType_loadToBD.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("LOAD_MARK_LINE_TO_DB ", true, false, false);
                })
            };
            panelSourceOptions.Items.Add(bt_addLineTypeToBD);


            var bt_addGroups = new RibbonButton
            {
                Name = "",
                Text = "Группы",
                ShowText = true,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/groups.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("GROUPS_TSODD ", true, false, false);
                })
            };
            panelSourceOptions.Items.Add(bt_addGroups);


            //* ************************************************          РЕДАКТИРОВАНИЕ            ************************************************ */

            RibbonPanelSource panelSourceSelection = new RibbonPanelSource
            {
                Title = "Выбор и редактирование"
            };

            RibbonPanel panel_6 = new RibbonPanel
            {
                Source = panelSourceSelection
            };
            tab.Panels.Add(panel_6);


            // кнопка перепривязки
            RibbonButton objectsBinding = new RibbonButton
            {
                Text = "Привязать объекты",
                ShowText = true,
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/stand_bindToAxis.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("BIND_OBJECTS ", true, false, false);

                })
            };
            panelSourceSelection.Items.Add(objectsBinding);


            var bt_quickSelection = new RibbonButton
            {
                Name = "",
                Text = "Выбор объектов",
                ShowText = true,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/selectionObjects.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("SELECT_TSODD_OBJECTS ", true, false, false);
                })
            };
            panelSourceSelection.Items.Add(bt_quickSelection);

            var bt_mleader = new RibbonButton
            {
                Name = "",
                Text = "Выноска",
                ShowText = true,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/mleader.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("MULTILEADER_TSODD ", true, false, false);
                })
            };
            panelSourceSelection.Items.Add(bt_mleader);

            quickProperties.Name = "";
            quickProperties.Text = "Свойства";
            quickProperties.ShowText = true;
            quickProperties.LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/quickProperties_OFF.png");
            quickProperties.Orientation = Orientation.Vertical;
            quickProperties.Size = RibbonItemSize.Large;
            quickProperties.CommandHandler = new RelayCommandHandler(() =>
            {
                TsoddHost.Current.doc?.SendStringToExecute("QUICK_PROPERTIES_TSODD_ON/OFF ", true, false, false);
            });
            panelSourceSelection.Items.Add(quickProperties);



            ///* ************************************************          Экспорт            ************************************************ */

            RibbonPanelSource panelSourceExport = new RibbonPanelSource
            {
                Title = "Экспорт ведомостей"
            };

            RibbonPanel panel_7 = new RibbonPanel
            {
                Source = panelSourceExport
            };
            tab.Panels.Add(panel_7);


            var bt_exportExcelSigns = new RibbonButton
            {
                Text = "Знаки",
                ShowText = true,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/xls_export_signs.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("EXPORT_SIGNS ", true, false, false);
                })
            };

            panelSourceExport.Items.Add(bt_exportExcelSigns);

            var bt_exportExcelMarks = new RibbonButton
            {
                Text = "Разметка",
                ShowText = true,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/xls_export_marks.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddHost.Current.doc?.SendStringToExecute("EXPORT_МАRKS ", true, false, false);
                })
            };

            panelSourceExport.Items.Add(bt_exportExcelMarks);


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
