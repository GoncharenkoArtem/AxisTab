using AxisTAb;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AxisTab;


public static class AxisTabCommands
{

    [CommandMethod("IA_NEW_AXIS")]
    [CommandMethod("IA_НОВАЯ_ОСЬ")]
    public static void Cmd_NewAxis()
    {
        Axis newAxis = new Axis();
        newAxis.NewAxis();
    }


    [CommandMethod("IA_SET_AXIS_NAME")]
    [CommandMethod("IA_ПЕРЕИМЕНОВАТЬ_ОСЬ")]
    public static void Cmd_AxisName()
    {
        var selectedAxis = RibbonInitializer.Instance.SelectAxis();
        if (selectedAxis == null) return;
        selectedAxis.GetAxisName();

        // перестраиваем combobox наименование осей на ribbon
        RibbonInitializer.Instance?.ListOFAxisRebuild();
    }

    [CommandMethod("IA_SET_AXIS_START_POINT")]
    [CommandMethod("IA_СТАРТОВАЯ_ТОЧКА_ОСИ")]
    public static void Cmd_AxisStartPoint()
    {
        var selectedAxis = RibbonInitializer.Instance.SelectAxis();
        if (selectedAxis == null) return;
        selectedAxis.GetAxisStartPoint();

        // перестраиваем combobox наименование осей на ribbon
        RibbonInitializer.Instance?.ListOFAxisRebuild();
    }

    [CommandMethod("IA_SET_PK")]
    [CommandMethod("IA_НАЗНАЧИТЬ_ПК")]
    public static void Cmd_SetPK()
    {
        RibbonInitializer.Instance.SetPkOnAxis();
    }

    [CommandMethod("IA_GET_PK")]
    [CommandMethod("IA_ПОЛУЧИТЬ_ПК")]
    public static void Cmd_GetPK()
    {
        RibbonInitializer.Instance.GetPkOnAxis();
    }


    [CommandMethod("IA_OPTIONS_AXIS")]
    [CommandMethod("IA_НАСТРОЙКИ_ОСИ")]
    public static void Cmd_OptionsTSODD()
    {
        OptionsForm optionsForm = new OptionsForm();
        optionsForm.ShowDialog();
    }


    [CommandMethod("IA_COCK_ANIMATION")]
    public static void Cmd_CockAnimation()
    {
      AnimationForm animForm = new AnimationForm();
      animForm.Show();
    }




    // вывод сообщения в editor
    private static void EditorMessage(string txt)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        ed.WriteMessage($" \n {txt}");
    }



}
