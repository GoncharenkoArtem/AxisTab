using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace ACAD_test
{
    // класс для записи/перезаписи и чтения XData полилиний
    public static class AutocadXData
    {
        // имя приложения, необходимо для регистрации XData в БД
        public const string AppName = "ITERIS_TSODD_APP";

        // метод, гарантирующий регистрацию приложения
        private static void GetAppReg(Transaction tr, Database db)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
           
            RegAppTable rat = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);   // таблица зарегистрированных приложений
            if (!rat.Has(AppName))
            {
                rat.UpgradeOpen();
                var rec = new RegAppTableRecord { Name = AppName };
                rat.Add(rec);
                tr.AddNewlyCreatedDBObject(rec, true);
            }
            
        }


        // метод записывет XData в Entity
        public static void UpdateXData(ObjectId objId, string axisHandle, string lineHandle = null)
        {
          
            if (objId.IsNull) throw new ArgumentNullException(nameof(objId));
            var db = objId.Database ?? Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database; // получаем БД объекта

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.GetDocument(db); // документ в котором объект
            if (doc != null)
            {
                using (doc.LockDocument())
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        // получаем Entity объект куда хотим записать XData
                        Entity ent = (Entity)tr.GetObject(objId, OpenMode.ForWrite);

                        var entDb = ent.Database; // та же база, но явно

                        // проверяем регистрацию приложения
                        GetAppReg(tr, entDb);

                        // создаем буфер для записи XData
                        ResultBuffer buff = new ResultBuffer {
                            new TypedValue ((int)DxfCode.ExtendedDataRegAppName, AppName),
                            new TypedValue ((int)DxfCode.ExtendedDataHandle, axisHandle)
                            };

                        // если есть Handle привязки к другой линии (для мультилинии)
                        if (lineHandle != null)
                        {
                            buff.Add(new TypedValue((int)DxfCode.ExtendedDataHandle, lineHandle));
                        }

                        // обновляем XData
                        ent.XData = AddXDataSection(ent.XData, buff);

                        tr.Commit();
                    }
                }
            }
        }



        // метод добавляет / обновляет секцию XData, не трогая старые секции
        private static ResultBuffer AddXDataSection(ResultBuffer buff, ResultBuffer newSection)
        {
            var outList = new List<TypedValue>();

            if (buff == null)
                return new ResultBuffer(newSection.AsArray());

            var arr = buff.AsArray();   // из ResultBuffer в TypedValue[] arr

            for (int i = 0; i < arr.Length;)
            {
                if (arr[i].TypeCode == (int)DxfCode.ExtendedDataRegAppName &&
                    string.Equals(arr[i].Value.ToString(), AppName, StringComparison.Ordinal))  // если это наше приложение 
                {
                    i++;    // мотаем до следующего приложения или конца XData массива
                    while (i < arr.Length && arr[i].TypeCode != (int)DxfCode.ExtendedDataRegAppName) i++;
                }
                else
                {
                    // добавляем в итоговый список текущий элемент TypedValue
                    outList.Add(arr[i]);
                    i++;
                }
            }

            outList.AddRange(newSection.AsArray());  // добавляем новую секцию

            return new ResultBuffer(outList.ToArray());  // отдаем перелапаченый ResultBuffer
        }


        // метод считывает XData у объекта 
        public static List<string> ReadXData(ObjectId objId)
        {
            List<string> result = new List<string>();

            if (objId.IsNull) throw new ArgumentNullException(nameof(objId));
            var db = objId.Database ?? Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database; // получаем БД объекта

            using (var tr = db.TransactionManager.StartTransaction())
            {
                Entity ent = (Entity)tr.GetObject(objId, OpenMode.ForRead);
                ResultBuffer buff = ent.XData;

                if (buff ==null) return result;

                var arr = buff.AsArray();

                for (int i = 0; i < arr.Length; i++)
                {
                    if (arr[i].TypeCode == (int)DxfCode.ExtendedDataRegAppName && arr[i].Value.ToString() == AppName)
                    {
                        i++;
                        while (i < arr.Length && arr[i].TypeCode != (int)DxfCode.ExtendedDataRegAppName)
                        { result.Add(arr[i].Value.ToString()); i++; }    // Добавляем XData в List

                        break;
                    }
                }
            }
            return result;
        }

    }
}