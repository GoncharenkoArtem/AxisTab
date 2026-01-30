using TSODD;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.IO.Pipelines;
using System.Linq;
using TSODD;
using TSODD.forms;
using TSODD.Forms;

public static class TsoddCommands
{

    [CommandMethod("NEW_AXIS")]
    [CommandMethod("НОВАЯ_ОСЬ")]
    public static void Cmd_NewAxis()
    {
        RibbonInitializer.Instance?.NewAxis();
    }

    [CommandMethod("SET_AXIS_NAME")]
    [CommandMethod("ПЕРЕИМЕНОВАТЬ_ОСЬ")]
    public static void Cmd_AxisName()
    {
        var selectedAxis = RibbonInitializer.Instance.SelectAxis();
        if (selectedAxis == null) return;
        selectedAxis.GetAxisName();

        // перестраиваем combobox наименование осей на ribbon
        RibbonInitializer.Instance?.ListOFAxisRebuild();
    }

    [CommandMethod("SET_AXIS_START_POINT")]
    [CommandMethod("СТАРТОВАЯ_ТОЧКА_ОСИ")]
    public static void Cmd_AxisStartPoint()
    {
        var selectedAxis = RibbonInitializer.Instance.SelectAxis();
        if (selectedAxis == null) return;
        selectedAxis.GetAxisStartPoint();

        // перестраиваем combobox наименование осей на ribbon
        RibbonInitializer.Instance?.ListOFAxisRebuild();
    }

    [CommandMethod("SET_PK")]
    [CommandMethod("НАЗНАЧИТЬ_ПК")]
    public static void Cmd_SetPK()
    {
        RibbonInitializer.Instance.SetPkOnAxis();
    }

    [CommandMethod("GET_PK")]
    [CommandMethod("ПОЛУЧИТЬ_ПК")]
    public static void Cmd_GetPK()
    {
        RibbonInitializer.Instance.GetPkOnAxis();
    }

    [CommandMethod("LAST_STAND_INSERT")]
    [CommandMethod("ПОСЛЕДНЯЯ_СТОЙКА")]
    public static void Cmd_InsertLastStand()
    {
        if (string.IsNullOrEmpty(TsoddHost.Current.currentStandBlock))
        {
            EditorMessage("\n Не определено имя последнего вставленного блока стойки \n");
            return;
        }
        TsoddBlock.InsertStandOrMarkBlock(TsoddHost.Current.currentStandBlock, true);
    }

    [CommandMethod("LAST_MARK_INSERT")]
    [CommandMethod("ПОСЛЕДНЯЯ_РАЗМЕТКА")]
    public static void Cmd_InsertLastMark()
    {
        if (string.IsNullOrEmpty(TsoddHost.Current.currentMarkBlock))
        {
            EditorMessage("\n Не определено имя последнего вставленного блока разметки \n");
            return;
        }
        TsoddBlock.InsertStandOrMarkBlock(TsoddHost.Current.currentMarkBlock, false);
    }

    [CommandMethod("LAST_SIGN_INSERT")]
    [CommandMethod("ПОСЛЕДНИЙ_ЗНАК")]
    public static void Cmd_InsertLastSign()
    {
        if (string.IsNullOrEmpty(TsoddHost.Current.currentSignBlock))
        {
            EditorMessage("\n Не определено имя последнего вставленного блока знака \n");
            return;
        }
        TsoddBlock.InsertSignBlock(TsoddHost.Current.currentSignBlock);
    }

    [CommandMethod("BIND_TO_AXIS")]
    [CommandMethod("ПРИВЯЗАТЬ_К_ОСИ")]
    public static void Cmd_BindToAxis()
    {
        TsoddBlock.ReBindStandBlockToAxis();
    }

    [CommandMethod("INSERT_TSODD_BLOCK")]
    [CommandMethod("ВСТАВИТЬ_БЛОК_ТСОДД")]
    public static void Cmd_Insert_TSODD_Block()
    {
        InsertBlockForm insertBlockForm = new InsertBlockForm();
        insertBlockForm.ShowDialog();
    }

    [CommandMethod("USER_MARK_BLOCK")]
    [CommandMethod("ПОЛЬЗОВАТЕЛЬСКИЙ_БЛОК_РАЗМЕТКИ")]
    public static void Cmd_UserMarkBlockk()
    {
        TsoddBlock.CreateUserMarkBlock();
    }

    [CommandMethod("MARK_LINE_INVERT")]
    [CommandMethod("ИНВЕРТИРОВАТЬ_ЛИНИЮ_РАЗМЕТКИ")]
    public static void Cmd_InvertLineType()
    {
        RibbonInitializer.Instance.LineTypeInvert();
    }

    [CommandMethod("MARK_LINE_TEXT_INVERT")]
    [CommandMethod("ПЕРЕСТАВИТЬ_ТЕКСТ_РАЗМЕТКИ")]
    public static void Cmd_LineTypeTextInvert()
    {
        RibbonInitializer.Instance.LineTypeTextInvert();
    }

    [CommandMethod("OPTIONS_TSODD")]
    [CommandMethod("НАСТРОЙКИ_ТСОДД")]
    public static void Cmd_OptionsTSODD()
    {
        OptionsForm optionsForm = new OptionsForm();
        optionsForm.ShowDialog();
    }

    [CommandMethod("LOAD_BLOCK_TO_DB")]
    [CommandMethod("ЗАГРУЗИТЬ_БЛОК_В_БД")]
    public static void Cmd_LoadBlock()
    {
        LoadBlock loadBlockForm = new LoadBlock();
        loadBlockForm.ShowDialog();
    }

    [CommandMethod("LOAD_MARK_LINE_TO_DB")]
    [CommandMethod("СОЗДАТЬ_ЛИНИЮ_РАЗМЕТКИ")]
    public static void Cmd_LoadLineTypeToDB()
    {
        LineTypeForm lineTypeForm = new LineTypeForm();
        lineTypeForm.ShowDialog();
    }

    [CommandMethod("GROUPS_TSODD")]
    [CommandMethod("ГРУППЫ_ТСОДД")]
    public static void Cmd_Groups()
    {
        GroupsAddForm groupsAddForm = new GroupsAddForm();
        groupsAddForm.ShowDialog();
    }

    [CommandMethod("SELECT_TSODD_OBJECTS")]
    [CommandMethod("ВЫБОР_ОБЪЕКТОВ_ТСОДД")]
    public static void Cmd_Selection()
    {
        ObjectSelectionForm objectSelectionForm = new ObjectSelectionForm();
        objectSelectionForm.ShowDialog();
    }

    [CommandMethod("MULTILEADER_TSODD")]
    [CommandMethod("МУЛЬТИВЫНОСКА_ТСОДД")]
    public static void Cmd_MultiLeader()
    {
        RibbonInitializer.Instance.CreateMLeaderForTsoddObject();
    }

    [CommandMethod("QUICK_PROPERTIES_TSODD_ON/OFF")]
    [CommandMethod("БЫСТРЫЕ_СВОЙСТВА_ТСОДД_ВКЛ/ВЫКЛ")]
    public static void Cmd_QuickProperties()
    {
        RibbonInitializer.Instance.quickPropertiesOn = !RibbonInitializer.Instance.quickPropertiesOn;
        if (RibbonInitializer.Instance.quickPropertiesOn) { RibbonInitializer.Instance.quickProperties.LargeImage = RibbonInitializer.Instance.LoadImage("pack://application:,,,/TSODD;component/images/quickProperties_ON.png"); }
        else { RibbonInitializer.Instance.quickProperties.LargeImage = RibbonInitializer.Instance.LoadImage("pack://application:,,,/TSODD;component/images/quickProperties_OFF.png"); }
    }

    [CommandMethod("EXPORT_SIGNS")]
    [CommandMethod("ВЕДОМОСТЬ_ЗНАКОВ")]
    public static void Cmd_ExportSigns()
    {
        var export = new ExportExcel();
        export.ExportSigns();
    }

    [CommandMethod("EXPORT_МАRKS")]
    [CommandMethod("ВЕДОМОСТЬ_РАЗМЕТКИ")]
    public static void Cmd_ExportMarks()
    {
        var export = new ExportExcel();
        export.ExportMarks();
    }

    [CommandMethod("COCK_ANIMATION")]
    public static void Cmd_CockAnimation()
    {
        ExportEnd exportEnd = new ExportEnd();
        exportEnd.ShowDialog();
    }





    // вывод сообщения в editor
    private static void EditorMessage(string txt)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        ed.WriteMessage($" \n {txt}");
    }



}
