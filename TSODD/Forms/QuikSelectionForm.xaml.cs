using TSODD;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace TSODD.Forms
{
    /// <summary>
    /// Логика взаимодействия для ObjectSelectionForm.xaml
    /// </summary>
    public partial class ObjectSelectionForm : Window
    {
        public ObjectSelectionForm()
        {
            InitializeComponent();

            var doc = TsoddHost.Current.doc;
            var locationX = doc.Window.DeviceIndependentLocation.X;
            var screenWidth = SystemParameters.PrimaryScreenWidth;

            this.Left = locationX + screenWidth / 2 - 150;
            this.Top = SystemParameters.PrimaryScreenHeight / 2 - 50;

            cb_TypeOfObject.SelectedIndex = 0;
            cb_SelectionType.SelectedIndex = 0;
        }

        private void cb_TypeOfObject_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cb_TypeOfObject.SelectedIndex == -1) return;
            cb_SelectionType.Items.Clear();

            switch (cb_TypeOfObject.SelectedIndex)
            {
                case 0: // оси
                    cb_SelectionType.Items.Add("весь чертеж");

                    break;

                case 1: // стойки
                    cb_SelectionType.Items.Add("весь чертеж");
                    cb_SelectionType.Items.Add("привязанные к оси");
                    break;

                case 2: // знаки
                    cb_SelectionType.Items.Add("весь чертеж");
                    cb_SelectionType.Items.Add("привязанные к оси");
                    cb_SelectionType.Items.Add("привязанные к стойке");
                    break;


                case 3:
                case 4:  // разметка линии,  разметка блоки
                    cb_SelectionType.Items.Add("весь чертеж");
                    cb_SelectionType.Items.Add("привязанные к оси");
                    break;
            }
            cb_SelectionType.SelectedIndex = 0;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var doc = TsoddHost.Current.doc;
            var db = doc.Database;
            var ed = doc.Editor;

            List<Stand> stands = TsoddBlock.GetListOfRefBlocks<Stand>();
            List<Sign> signs = TsoddBlock.GetListOfRefBlocks<Sign>();
            List<Mark> marks = TsoddBlock.GetListOfRefBlocks<Mark>();

            switch (cb_TypeOfObject.SelectedIndex.ToString() + cb_SelectionType.SelectedIndex.ToString())
            {
                case "00":  // оси весь чертеж

                    SelectAllAxis();
                    break;

                case "10":  // стойки весь чертеж

                    SelectObjects(selectedIds<Stand>(stands));
                    break;

                case "11":  // стойки привязанные к оси

                    var curAxis = RibbonInitializer.Instance?.SelectAxis();
                    if (curAxis == null)
                    {
                        ed.WriteMessage("\n Ошибка выбора оси \n");
                        this.Close();
                        return;
                    }

                    // фильтруем стойки по оси 
                    stands = stands.Where(s => s.AxisName == curAxis.Name).ToList();
                    SelectObjects(selectedIds<Stand>(stands));
                    break;

                case "20":  // знаки весь чертеж

                    SelectObjects(selectedIds<Sign>(signs));
                    break;

                case "21": // знаки, привязанные к оси

                    curAxis = RibbonInitializer.Instance?.SelectAxis();
                    if (curAxis == null)
                    {
                        ed.WriteMessage("\n Ошибка выбора оси \n");
                        this.Close();
                        return;
                    }

                    // фильтруем знаки по оси 
                    List<Sign> separatedSigns = new List<Sign>();
                    foreach (var sign in signs)
                    {
                        Stand st = stands.FirstOrDefault(s => s.Handle == sign.StandHandle);
                        if (st == null) continue;
                        if (st.AxisName == curAxis.Name) separatedSigns.Add(sign);
                    }
                    SelectObjects(selectedIds<Sign>(separatedSigns));
                    break;


                case "22": // знаки, привязанные к стойке

                    PromptSelectionOptions pso = new PromptSelectionOptions();
                    pso.MessageForAdding = "\n Выберите блок стойки ";
                    pso.SingleOnly = true;

                    var filter = new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "INSERT") });

                    var psr = ed.GetSelection(pso, filter);
                    if (psr.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\n Ошибка выбора блока стойки \n");
                        this.Close();
                        return;
                    }

                    string standHandle = psr.Value[0].ObjectId.Handle.ToString();

                    // фильтруем знаки по блоку стойки
                    signs = signs.Where(s => s.StandHandle == standHandle).ToList();

                    SelectObjects(selectedIds<Sign>(signs));
                    break;

                case "30":  // все типы линий

                    var objectIds = LineTypeReader.CollectMarkLineTypeID(onlyMaster: false);
                    SelectObjects(objectIds);
                    break;

                case "31":  // типы линии, привязянные к оси

                    curAxis = RibbonInitializer.Instance?.SelectAxis();
                    if (curAxis == null)
                    {
                        ed.WriteMessage("\n Ошибка выбора оси \n");
                        this.Close();
                        return;
                    }

                    List<ObjectId> separetedIds = new List<ObjectId>();

                    objectIds = LineTypeReader.CollectMarkLineTypeID(onlyMaster: false);
                    foreach (var id in objectIds)
                    {
                        TsoddXdataElement tsoddXdataElement = new TsoddXdataElement();
                        tsoddXdataElement.Parse(id);

                        if (tsoddXdataElement.AxisHandle == curAxis.PolyHandle.ToString())
                        {
                            separetedIds.Add(id);
                            if (tsoddXdataElement.SlavePolylineID != ObjectId.Null) separetedIds.Add(tsoddXdataElement.SlavePolylineID);
                            if (tsoddXdataElement.MtextID != ObjectId.Null) separetedIds.Add(tsoddXdataElement.MtextID);
                        }
                    }

                    objectIds = new ObjectId[separetedIds.Count];
                    for (int i = 0; i < separetedIds.Count; i++) objectIds[i] = separetedIds[i];

                    SelectObjects(objectIds);

                    break;

                case "40":  // все блоки разметки

                    SelectObjects(selectedIds<Mark>(marks));
                    break;

                case "41":  // все блоки разметки

                    curAxis = RibbonInitializer.Instance?.SelectAxis();
                    if (curAxis == null)
                    {
                        ed.WriteMessage("\n Ошибка выбора оси \n");
                        this.Close();
                        return;
                    }

                    // фильтруем разметку по оси 
                    var separatedMarks = marks.Where(m => m.AxisName == curAxis.Name).ToList();
                    SelectObjects(selectedIds<Mark>(separatedMarks));
                    break;
            }

            // внутренний метод конвертации списка в массив ObjectIds
            ObjectId[] selectedIds<T>(List<T> listIds) where T : IRefBlock
            {
                ObjectId[] objectIds = new ObjectId[listIds.Count];
                for (int i = 0; i < listIds.Count; i++) objectIds[i] = listIds[i].ID;
                return objectIds;
            }

            this.Close();
        }



        private void SelectAllAxis()
        {
            var doc = TsoddHost.Current.doc;
            var db = doc.Database;
            var ed = doc.Editor;

            var filter = new SelectionFilter(new TypedValue[]
            { new TypedValue((int)DxfCode.Start,"LWPOLYLINE,POLYLINE"),
                new TypedValue((int)DxfCode.ExtendedDataRegAppName,AutocadXData.AppName),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString,"axis")
            }
            );

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // поиск с учетом фильтра
                var selection = ed.SelectAll(filter); ;

                if (selection.Status != PromptStatus.OK)    // неудачный поиск
                {
                    ed.WriteMessage("\n В чертеже не найдено осей \n");
                    return;
                }
                else
                {
                    ed.SetImpliedSelection(selection.Value);
                    ed.WriteMessage($"\n Выбрано {selection.Value.Count} ось(и) \n");
                    ed.UpdateScreen();
                }
            }
        }

        //private Axis SelectAxis()
        //{
        //    var doc = TsoddHost.Current.doc;
        //    var db = doc.Database;
        //    var ed = doc.Editor;

        //   var filter = new SelectionFilter(new TypedValue[]
        //          { new TypedValue((int)DxfCode.Start,"LWPOLYLINE,POLYLINE"),
        //                            new TypedValue((int)DxfCode.ExtendedDataRegAppName,AutocadXData.AppName),
        //                            new TypedValue((int)DxfCode.ExtendedDataAsciiString,"axis")
        //          });

        //    // Id выбранной оси
        //    var axisId = RibbonInitializer.Instance.GetAutoCadSelectionObjectsId(filter, true);
        //    if (axisId == null)
        //    {
        //        ed.WriteMessage("\n Ошибка выбора оси \n");
        //        return null;
        //    }

        //    // получаем сам объект оси
        //    return TsoddHost.Current.axis.FirstOrDefault(a => a.PolyID == axisId[0]);
        //}


        //private ObjectId[] SelectMarkLineType(bool onlyMaster)
        //{
        //    var doc = TsoddHost.Current.doc;
        //    var db = doc.Database;
        //    var ed = doc.Editor;

        //    SelectionFilter filter;
        //    if (onlyMaster)
        //    {
        //      filter = new SelectionFilter(new TypedValue[]
        //            { 
        //              new TypedValue((int)DxfCode.ExtendedDataRegAppName,AutocadXData.AppName),
        //              new TypedValue((int)DxfCode.ExtendedDataAsciiString,"master")
        //            });
        //    }
        //    else
        //    {
        //        filter = new SelectionFilter(new TypedValue[]
        //              { 
        //              new TypedValue((int)DxfCode.ExtendedDataRegAppName,AutocadXData.AppName),
        //              new TypedValue((int)DxfCode.ExtendedDataAsciiString,"master,slave")
        //              });
        //    }

        //    using (doc.LockDocument())
        //    using (var tr = db.TransactionManager.StartTransaction())
        //    {
        //        // поиск с учетом фильтра
        //        var selection = ed.SelectAll(filter);

        //        if (selection.Status != PromptStatus.OK)    // неудачный поиск
        //        {
        //            ed.WriteMessage("\n В чертеже не найдено разметки в виде линий \n");
        //            return new ObjectId[0];
        //        }
        //        else
        //        {
        //            return selection.Value.GetObjectIds();
        //        }
        //    }

        //}



        private void SelectObjects(ObjectId[] ids)
        {
            var doc = TsoddHost.Current.doc;
            var db = doc.Database;
            var ed = doc.Editor;

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                if (ids.Length == 0)    // неудачный поиск
                {
                    ed.WriteMessage("\n Ошибка выбора объектов. Нет подходящих объектов под условия поиска \n");
                    ed.SetImpliedSelection(new ObjectId[0]);
                    return;
                }
                else
                {
                    ed.SetImpliedSelection(ids);
                    ed.WriteMessage($"\n Выбрано {ids.Length} объект(а) \n");
                    ed.UpdateScreen();
                }
            }


        }






    }
}
