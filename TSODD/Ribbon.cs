using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System.Windows.Input;
using System.Linq;
using System;
using System.Windows;
using System.Windows.Media.Imaging;
using TSODD;
using Autodesk.AutoCAD.ApplicationServices;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using System.Windows.Controls;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Windows.Data;
using Autodesk.AutoCAD.EditorInput;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Autodesk.AutoCAD.Windows;
using System.Xml.Linq;


namespace ACAD_test
{
    public partial class RibbonInitializer : IExtensionApplication
    {
        public static RibbonInitializer Instance { get; private set; }
        private RibbonCombo axisCombo = new RibbonCombo();
        public RibbonCombo signsGroups = new RibbonCombo();
        public RibbonCombo marksCombo = new RibbonCombo();

        private RibbonSplitButton splitStands = new RibbonSplitButton();
        private RibbonSplitButton splitSigns = new RibbonSplitButton();
        public RibbonSplitButton splitMarksLineTypes = new RibbonSplitButton();

        public RibbonRowPanel rowLineType = new RibbonRowPanel();
        private RibbonCombo comboLineTypePattern_1 = new RibbonCombo();
        private RibbonCombo comboLineTypePattern_2 = new RibbonCombo();
        private RibbonCombo comboLineTypeWidth_1 = new RibbonCombo();
        private RibbonCombo comboLineTypeWidth_2 = new RibbonCombo();
        private RibbonCombo comboLineTypeColor_1 = new RibbonCombo();
        private RibbonCombo comboLineTypeColor_2 = new RibbonCombo();
        private RibbonCombo comboLineTypeOffset = new RibbonCombo();
        private RibbonLabel labelLineType = new RibbonLabel(); 


        //  private RibbonSplitButton splitMarks = new RibbonSplitButton();


        public RibbonPanelSource panelSourceMarks;



        private static readonly HashSet<IntPtr> _dbIntPtr = new HashSet<IntPtr>();
        private readonly HashSet<ObjectId> _dontDeleteMe = new HashSet<ObjectId>();
        private readonly HashSet<ObjectId> _deleteMe = new HashSet<ObjectId>();

        private bool _marksLineTypeFlag = true;


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
                splitSigns.CurrentChanged += SplitSigns_CurrentChanged;
                signsGroups.CurrentChanged += SignsGroups_CurrentChanged;

                AddRibbonPanel();

                // предвыбор значений выпадающих списков
                // группа знаков
                if (signsGroups.Items.Count > 0)
                {
                    var currSignGroup = signsGroups.Items[0] as RibbonButton;
                    TsoddHost.Current.currentSignGroup = currSignGroup.Text ;
                }
                // стойка
                if (splitStands.Items.Count > 0) TsoddHost.Current.currentStandBlock = splitStands.Items[0].Text;
                // знак
                if (splitSigns.Items.Count > 0) TsoddHost.Current.currentSignBlock = splitSigns.Items[0].Text;

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
                FillBlocksMenu(splitStands, "STAND");
                // настраиваем контрол со знаками
                InitializeSplitButtons(splitSigns, "Знаки");
                // заполняем 
                FillBlocksMenu(splitSigns, "SIGN", TsoddHost.Current.currentSignGroup);
                // обновляем  элементы LineType
                ListOfMarksLinesLoad(200, 20);
                LineTypeReader.RefreshLineTypesInAcad();
            }
        }

   

        private void AddRibbonPanel()
        {
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
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/i20.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,
                CommandHandler = new RelayCommandHandler(() =>
                {
                    NewAxis();
                })
            };
            panelSourceAxis.Items.Add(bt_newAxis);
            panelSourceAxis.Items.Add(new RibbonSeparator());

            var rowAxis = new RibbonRowPanel();

            // список combobox с осями 
            axisCombo.Text = " Текущая ось ";
            axisCombo.ShowText = true;
            axisCombo.Width = 210;
            rowAxis.Items.Add(axisCombo);

            // Перенос на новую строку внутри этой же группы
            rowAxis.Items.Add(new RibbonRowBreak());

            //  Кнопка изменения имени оси
            var bt_changeName = new RibbonButton
            {
                Text = "Редактировать имя оси",
                ShowText = true,
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
                Image = LoadImage("pack://application:,,,/TSODD;component/images/i20_small.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddCommands.Cmd_AxisName();
                })
            };

    
            // Кнопка изменения начальной точки
            var bt_startPoint = new RibbonButton
            {
                Text = "Редактировать начальную точку ",
                ShowText = true,
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
                Image = LoadImage("pack://application:,,,/TSODD;component/images/i20_small.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddCommands.Cmd_AxisStartPoint();
                })
            };

            // Кнопки рядом + разделитель между ними
            rowAxis.Items.Add(bt_changeName);

            rowAxis.Items.Add(new RibbonRowBreak());

            rowAxis.Items.Add(bt_startPoint);

            // добавляем группу "ось" в панель
            panelSourceAxis.Items.Add(rowAxis);

        
            /* ************************************************           СТОЙКИ            ************************************************ */

            RibbonPanelSource panelSourceStands = new RibbonPanelSource
            {
                Title = "Стойки"
            };

            RibbonPanel panel_2 = new RibbonPanel
            {
                Source = panelSourceStands
            };
            tab.Panels.Add(panel_2);

      
            // добавляем на ribbon
            panelSourceStands.Items.Add(splitStands);

            // кнопки для стоек
            var rowStandButtons = new RibbonRowPanel();

            // кнопка перепривязки стоек к оси
            RibbonButton axisBinding = new RibbonButton
            {
                Text = "Привязка к оси",
                ShowText = true,
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
                Image = LoadImage("pack://application:,,,/TSODD;component/images/i20_small.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddBlock.ReBindStandBlockToAxis();
          
                })
            };
            rowStandButtons.Items.Add(axisBinding);
            rowStandButtons.Items.Add(new RibbonRowBreak());

            // добавить блок стойки в базу
            RibbonButton loadBlockToBD = new RibbonButton {
                Text = "Добавть стойку",
                ShowText = true,
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
                Image = LoadImage("pack://application:,,,/TSODD;component/images/i20_small.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    // подгрузка блоков в БД
                    string blockName = TsoddBlock.AddBlockToBD("STAND_TEMPLATE");
                    // пересобираем список
                    splitStands.Items.Clear();
                    FillBlocksMenu(splitStands, "STAND", blockName);

                })
            };
            rowStandButtons.Items.Add(loadBlockToBD);
            rowStandButtons.Items.Add(new RibbonRowBreak());

            // добавить блок в базу
            RibbonButton deleteStandFromBD = new RibbonButton
            {
                Text = "Удалить стойку",
                ShowText = true,
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
                Image = LoadImage("pack://application:,,,/TSODD;component/images/i20_small.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    // удаление блоков из БД
                    string blockName = String.Empty;
                    if (splitStands.Current != null)
                    {
                        blockName = splitStands.Current.Text;
                    }
                    else 
                    {
                        return; // имени блока нет, выходим
                    }

                    TsoddBlock.DeleteBlockFromBD(blockName);
                    
                    // пересобираем список
                    splitStands.Items.Clear();
                    FillBlocksMenu(splitStands, "STAND");

                    // первый элемент
                    var firstElement = splitStands.Items.FirstOrDefault();
                    if(firstElement!=null) splitStands.Current = firstElement;


                })
            };
            rowStandButtons.Items.Add(deleteStandFromBD);
            rowStandButtons.Items.Add(new RibbonRowBreak());

            panelSourceStands.Items.Add(rowStandButtons);

            /* ************************************************           ЗНАКИ            ************************************************ */

            RibbonPanelSource panelSourceSigns = new RibbonPanelSource
            {
                Title = "Знаки"
            };

            RibbonPanel panel_3 = new RibbonPanel
            {
                Source = panelSourceSigns
            };
            tab.Panels.Add(panel_3);

      
            var rowSignsGroup = new RibbonRowPanel();
            var rowSignsGroup_1 = new RibbonRowPanel();
            var rowSignsGroup_2 = new RibbonRowPanel();

            // список combobox с группами знаков
            signsGroups.Text = "Группа знаков ";
            signsGroups.ShowText = true;
            signsGroups.Width = 280;
            rowSignsGroup.Items.Add(signsGroups);


            // заполняем комбобокс с группами знаков
            TsoddBlock.FillSignsGroups();

            // Перенос на новую строку внутри этой же группы
            rowSignsGroup.Items.Add(new RibbonRowBreak());

            //  Кнопка добавления группы
            var bt_AddGroup = new RibbonButton
            {
                Text = "Добавить группу",
                ShowText = true,
                MinWidth = 140,
                
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
                Image = LoadImage("pack://application:,,,/TSODD;component/images/i20_small.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    TsoddBlock.AddSignGroupToBD();
                })
            };

            // Кнопка удаления группы
            var bt_DeleteGroup = new RibbonButton
            {
                Text = " Удалить группу",
                ShowText = true,
                MinWidth = 140,
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
                Image = LoadImage("pack://application:,,,/TSODD;component/images/i20_small.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {

                    TsoddBlock.DeleteSignGroupFromBD(TsoddHost.Current.currentSignGroup);
                })
            };

            // Кнопки рядом + разделитель между ними
            rowSignsGroup_1.Items.Add(bt_AddGroup);
            rowSignsGroup_1.Items.Add(new RibbonRowBreak());
            rowSignsGroup_1.Items.Add(bt_DeleteGroup);

            // добавить блок знака в БД
            var bt_addSignToBD = new RibbonButton
            {
                Text = "Добавить знак",
                ShowText = true,
                MinWidth = 140,
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
                Image = LoadImage("pack://application:,,,/TSODD;component/images/i20_small.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    // подгрузка блоков в БД
                    string blockName = TsoddBlock.AddBlockToBD("SIGN_TEMPLATE");
                    // пересобираем список
                    splitSigns.Items.Clear();
                    FillBlocksMenu(splitSigns, "SIGN", blockName);
                })
            };

            // Удалить блок знака из БД
            var bt_deleteSignFromBD = new RibbonButton
            {
                Text = "Удалить знак",
                ShowText = true,
                MinWidth = 140,
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
                Image = LoadImage("pack://application:,,,/TSODD;component/images/i20_small.png"),
                CommandHandler = new RelayCommandHandler(() =>
                {
                    // удаление блоков из БД
                    string blockName = String.Empty;
                    if (splitSigns.Current != null)
                    {
                        blockName = splitSigns.Current.Text;
                    }
                    else
                    {
                        return; // имени блока нет, выходим
                    }

                    TsoddBlock.DeleteBlockFromBD(blockName);

                    // пересобираем список
                    splitSigns.Items.Clear();
                    FillBlocksMenu(splitSigns, "SIGN", TsoddHost.Current.currentSignGroup);
                    splitSigns.ListStyle = RibbonSplitButtonListStyle.Descriptive;

                    // первый элемент
                    var firstElement = splitSigns.Items.FirstOrDefault();
                    if (firstElement != null) splitSigns.Current = firstElement;

                })
            };


            rowSignsGroup_2.Items.Add(bt_addSignToBD);
            rowSignsGroup_2.Items.Add(new RibbonRowBreak());
            rowSignsGroup_2.Items.Add(bt_deleteSignFromBD);

            rowSignsGroup.Items.Add(rowSignsGroup_1);
            rowSignsGroup.Items.Add(rowSignsGroup_2);

            // добавляем группу в панель
            panelSourceSigns.Items.Add(rowSignsGroup);

            // добавляем на ribbon
            panelSourceSigns.Items.Add(splitSigns);



            /* ************************************************          РАЗМЕТКА            ************************************************ */

            panelSourceMarks = new RibbonPanelSource
            {
                Title = "Разметка"
            };

            RibbonPanel panel_4 = new RibbonPanel
            {
                Source = panelSourceMarks
            };
            tab.Panels.Add(panel_4);

            splitMarksLineTypes.Width = 80;
            splitMarksLineTypes.Size = RibbonItemSize.Large;
            splitMarksLineTypes.IsSplit = true;
            panelSourceMarks.Items.Add(splitMarksLineTypes);

            // типы линий
            //rowLineType = new RibbonRowPanel();
            
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

            var bt_newB = new RibbonButton
            {
                Name = "",
                Text = "TEMP",
                ShowText = true,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/i20.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,
                CommandHandler = new RelayCommandHandler(() =>
                {
                   // ListOfMarksLinesLoad(200,20);
                   // LineTypeReader.Test();
                   
                })
            };

            panelSourceMarks.Items.Add(bt_newB);

            var bt_newC = new RibbonButton
            {
                Name = "",
                Text = "TEMP2",
                ShowText = true,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/i20.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,
                CommandHandler = new RelayCommandHandler(() =>
                {
                    //ListOfMarksLinesLoad(200,20);
                    //LineTypeReader.Test();
                    LineTypeReader.Test2();


                })
            };

            panelSourceMarks.Items.Add(bt_newC);

            /* ************************************************          Выбор            ************************************************ */

            RibbonPanelSource panelSourcePick = new RibbonPanelSource
            {
                Title = "Выбор"
            };

            RibbonPanel panel_5= new RibbonPanel

            {
                Source = panelSourcePick
            };
            tab.Panels.Add(panel_5);

            /* ************************************************          Экспорт            ************************************************ */

            RibbonPanelSource panelSourceExport = new RibbonPanelSource
            {
                Title = "Экспорт"
            };

            RibbonPanel panel_6 = new RibbonPanel
            {
                Source = panelSourceExport
            };
            tab.Panels.Add(panel_6);


            // Удалить блок знака из БД
            var bt_exportExcel = new RibbonButton
            {
                Text = "Эксопрт в Excel",
                ShowText = true,
                LargeImage = LoadImage("pack://application:,,,/TSODD;component/images/i20.png"),
                Orientation = Orientation.Vertical,
                Size = RibbonItemSize.Large,
                CommandHandler = new RelayCommandHandler(() =>
                {
                    var export = new ExportExcel();
                    export.CreateTableHeader();
                    export.ExportSigns();

                })
            };

            panelSourceExport.Items.Add(bt_exportExcel);












            //var rowStands = new RibbonRowPanel();


            //var splitStands = new RibbonSplitButton
            //{
            //    Text = "Стойки",
            //    Size = RibbonItemSize.Large,
            //    Orientation = Orientation.Vertical,
            //    IsSplit = true, // раскрывающееся меню
            //    ShowText = true,

            //    Width = 100


            //splitSigns
            //{
            //    Text = "Знаки",
            //    Size = RibbonItemSize.Large,
            //    Orientation = Orientation.Vertical,
            //    IsSplit = true, // раскрывающееся меню
            //                    //ShowText = true,
            //    MinHeight = 80,
            //    Height = 80,

            //    Width = 100
            //};


            //FillBlocksMenu(splitSigns, "SIGN");

            //panelSourceStands.Items.Add(splitSigns);


            //var newlist = GetListOfBlocks("SIGN");

            //// 1) Большая превью-кнопка (вставляет текущий блок)
            //RibbonButton _previewBtn = new RibbonButton
            //{
            //    Text = "(не выбрано)",
            //    ShowText = false,
            //    Size = RibbonItemSize.Large,
            //    Orientation = Orientation.Horizontal,
            //    LargeImage = newlist[0].img,
            //    Height = 100,
            //    Width = 100


            //};
            //_previewBtn.CommandHandler = new RelayCommandHandler(() =>
            //{
            //    //if (!string.IsNullOrEmpty(_currentBlockName))
            //    //    InsertBlockByName(_currentBlockName);
            //});
            //panelSourceStands.Items.Add(_previewBtn);

            //// перенос строки, чтобы маленькая была «под» большой (или не добавляй — будет справа)
            //panelSourceStands.Items.Add(new RibbonRowBreak());

            //// 2) Маленькая кнопка-список (вся кнопка открывает меню)
            //var ddBlocks = new RibbonSplitButton
            //{
            //    Text = "Выбрать…",         // можно без текста и с иконкой «стрелка»
            //    ShowText = true,
            //    Size = RibbonItemSize.Standard,
            //    Orientation = Orientation.Horizontal,
            //    IsSplit = false            // 👈 важное: нет «верх-низ», вся кнопка — дропдаун
            //};

            //FillBlocksMenu(ddBlocks, "SIGNS");
            //panelSourceStands.Items.Add(ddBlocks);

        }



















        //// допустим, у тебя уже есть panelSourceAxis и combo:
        //var blocksCombo = new RibbonCombo
        //{
        //    Text = "Блок",
        //    ShowText = true,
        //    Width = 220
        //};

        //// заполняем блоками из текущей базы
        //var db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;

        //using (var tr = db.TransactionManager.StartTransaction())
        //{
        //    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        //    foreach (ObjectId id in bt)
        //    {
        //        var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
        //        // пропустим служебные/анонимные/пространства
        //        if (btr.IsLayout || btr.IsAnonymous || btr.Name == BlockTableRecord.ModelSpace || btr.Name == BlockTableRecord.PaperSpace)
        //            continue;

        //        var bmp = TsoddBlock.GetBlockPreviewBitmap(btr);
        //        var img = TsoddBlock.ToImageSource(bmp);

        //        var item = new RibbonButton
        //        {
        //            Text = btr.Name,
        //            ShowText = true,
        //            ShowImage = img != null,
        //            //Image = img,                      // маленькая картинка в выпадающем списке
        //            LargeImage = img,                   // на всякий
        //            Size = RibbonItemSize.Standard,
        //            Orientation = System.Windows.Controls.Orientation.Horizontal,
        //            CommandParameter = btr.Name
        //        };

        //        item.CommandHandler = new RelayCommandHandler(() =>
        //        {
        //            // здесь твоя логика «вставить выбранный блок»
        //            var name = item.CommandParameter as string;
        //            //InsertBlockByName(name);
        //        });

        //        blocksCombo.Items.Add(item);

        //    }
        //    tr.Commit();
        //}

        //panelSourceAxis.Items.Add(blocksCombo);








        //// === ГРУППА "стойки" ===
        //var rowPosts = new RibbonRowPanel();

        //// 2.1 Комбобокс
        //var postsCombo = new RibbonCombo
        //{
        //    Text = "Стойка",
        //    ShowText = true,
        //    Width = 160
        //};
        //rowPosts.Items.Add(postsCombo);

        //// Разделитель между «колонками» внутри группы
        //rowPosts.Items.Add(new RibbonSeparator());

        //// 2.2 Три кнопки в столбик справа
        //var stack = new RibbonItemCollection();

        //var bt4 = new RibbonButton { Text = "bt4", ShowText = true, Size = RibbonItemSize.Standard };
        //bt4.CommandHandler = new RelayCommandHandler(() => { /* ... */
        //});

        //var bt5 = new RibbonButton { Text = "bt5", ShowText = true, Size = RibbonItemSize.Standard };
        //bt5.CommandHandler = new RelayCommandHandler(() => { /* ... */ });

        //var bt6 = new RibbonButton { Text = "bt6", ShowText = true, Size = RibbonItemSize.Standard };
        //bt6.CommandHandler = new RelayCommandHandler(() => { /* ... */ });

        //rowPosts.Items.Add(bt4);
        //rowPosts.Items.Add(bt5);
        //rowPosts.Items.Add(bt6);


        //// добавляем группу "стойки" в панель
        //panelSourceAxis.Items.Add(rowPosts);





        //// Кнопка "Новая ось"
        //RibbonButton button = new RibbonButton
        //{
        //    Name = "NEWAXIS",
        //    Text = "Новая ось",
        //    ShowText = true,
        //    //LargeImage = LoadImage("pack://application:,,,/ACAD_test;component/images/icon.png"),
        //    Orientation = System.Windows.Controls.Orientation.Vertical,
        //    Size = RibbonItemSize.Large,
        //    CommandHandler = new RelayCommandHandler(() =>
        //    {
        //       TsoddHost.Current.NewAxis();
        //    })
        //};

        //panelSourceAxis.Items.Add(button);




        //// Кнопка
        //RibbonButton button2 = new RibbonButton
        //{
        //    Name = "22",
        //    Text = "starpoint",
        //    ShowText = true,
        //    //LargeImage = LoadImage("pack://application:,,,/ACAD_test;component/images/icon.png"),
        //    Orientation = System.Windows.Controls.Orientation.Vertical,
        //    Size = RibbonItemSize.Standard,
        //    CommandHandler = new RelayCommandHandler(() =>
        //    {

        //List<string> ll = GetListOfBlocks("STAND");

        //    })
        //};
        //panelSourceAxis.Items.Add(button2);

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
