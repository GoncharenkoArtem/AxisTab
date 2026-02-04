using AxisTAb;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;


public static class AxisTabCommands
{

    [CommandMethod("NEW_AXIS_1")]
    [CommandMethod("НОВАЯ_ОСЬ_1")]
    public static void Cmd_NewAxis()
    {
        Axis newAxis = new Axis();
        newAxis.NewAxis();
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



    [CommandMethod("COCK_ANIMATION")]
    public static void Cmd_CockAnimation()
    {
        //ExportEnd exportEnd = new ExportEnd();
       // exportEnd.ShowDialog();
    }




    // вывод сообщения в editor
    private static void EditorMessage(string txt)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        ed.WriteMessage($" \n {txt}");
    }



}
