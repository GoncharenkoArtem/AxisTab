using ACAD_test;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Linq;
using TSODD;
using static System.Net.Mime.MediaTypeNames;

public static class TsoddCommands
{
    
    [CommandMethod("NewAxis")]
    public static void Cmd_NewAxis()
    {
        RibbonInitializer.Instance?.NewAxis(); 
    }


    [CommandMethod("SetAxisName")]
    public static void Cmd_AxisName()
    {
        var selectedAxis = SelectAxis();
        if (selectedAxis == null) return;
        selectedAxis.GetAxisName();
        
        // перестраиваем combobox наименование осей на ribbon
        RibbonInitializer.Instance?.ListOFAxisRebuild();
    }


    [CommandMethod("SetAxisStartPoint")]
    public static void Cmd_AxisStartPoint()
    {
        var selectedAxis = SelectAxis();
        if (selectedAxis == null) return;
        selectedAxis.GetAxisStartPoint();

        // перестраиваем combobox наименование осей на ribbon
        RibbonInitializer.Instance?.ListOFAxisRebuild();

    }

    // метод выбора и проверки полилинии оси
    public static Axis SelectAxis()
    {
        Axis resultAxis = null;
        var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        // Настройки промпта
        var peo = new PromptEntityOptions("\n Выберите ось: ");
        peo.SetRejectMessage("\n Это не полилиния. Выберите полилинию!");
        peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), exactMatch: false);

        // Запрос
        var per = ed.GetEntity(peo);
        if (per.Status == PromptStatus.OK)
        {
            resultAxis = TsoddHost.Current.axis.FirstOrDefault(a => a.PolyID == per.ObjectId);
            if (resultAxis == null)
            {
                EditorMessage("\n Данная полилиния не является осью ");
                return null;
            }
        }
       
        return resultAxis;
    }


    [CommandMethod("InsertStand")]
    public static void Cmd_InsertStand()
    {
       TsoddBlock.InsertStandOrMarkBlock(TsoddHost.Current.currentStandBlock);
    }



    [CommandMethod("InsertSign")]
    public static void Cmd_InsertSign()
    {
        TsoddBlock.InsertSignBlock(TsoddHost.Current.currentStandBlock);
    }





    // вывод сообщения в editor
    private static void EditorMessage(string txt)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        ed.WriteMessage($" \n {txt}");
    }



}
