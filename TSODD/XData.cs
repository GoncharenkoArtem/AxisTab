using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;





namespace TSODD
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
        public static void UpdateXData(ObjectId objId, List<(int code, string value)> list)
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
                        ResultBuffer buff = new ResultBuffer { new TypedValue((int)DxfCode.ExtendedDataRegAppName, AppName) };
                        if (list != null) foreach (var val in list) buff.Add(new TypedValue(val.code, val.value));

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
        public static List<(int, string)> ReadXData(ObjectId objId)
        {
            List<(int, string)> result = new List<(int, string)>();

            if (objId.IsNull) throw new ArgumentNullException(nameof(objId));
            var db = objId.Database ?? Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database; // получаем БД объекта

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    Entity ent = (Entity)tr.GetObject(objId, OpenMode.ForRead);

                    ResultBuffer buff = ent.XData;

                    if (buff == null) return result;

                    var arr = buff.AsArray();

                    for (int i = 0; i < arr.Length; i++)
                    {
                        if (arr[i].TypeCode == (int)DxfCode.ExtendedDataRegAppName && arr[i].Value.ToString() == AppName)
                        {
                            i++;
                            while (i < arr.Length && arr[i].TypeCode != (int)DxfCode.ExtendedDataRegAppName)
                            { result.Add((arr[i].TypeCode, arr[i].Value.ToString())); i++; }    // Добавляем XData в List

                            break;
                        }
                    }
                    tr.Commit();
                }
                return result;
            }
            catch (Exception ex) { return result; }
        }
    }





    // перечисление типов объектов Tsodd
    public enum TsoddElement { Line, Mline, Axis }

    /// <summary>
    ///  Класс для быстрой идентификации элементов Tsodd и парсинга их Xdata
    /// </summary>
    public class TsoddXdataElement
    {
        public TsoddElement Type { get; set; }
        public ObjectId MasterPolylineID { get; set; } = ObjectId.Null;
        public ObjectId SlavePolylineID { get; set; } = ObjectId.Null;
        public ObjectId MtextID { get; set; } = ObjectId.Null;
        public ObjectId AxisPolylineID { get; set; } = ObjectId.Null;
        public string AxisHandle { get; set; } = string.Empty;
        public string Number { get; set; } = string.Empty;
        public string Material { get; set; } = string.Empty;
        public string Existence { get; set; } = string.Empty;

        public List<(int code, string value)> ListXdata { get; set; } = new List<(int code, string value)>();

        public void Parse(ObjectId objectId)
        {
            ListXdata = AutocadXData.ReadXData(objectId);
            if (ListXdata.Count == 0) { return; } //  пустая XData

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            Handle tempHandle;
            string tempString;

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {

                    switch (ListXdata[0].value)
                    {
                        case "axis":
                            Type = TsoddElement.Axis;

                            tempHandle = new Handle(Convert.ToInt64(ListXdata[1].value, 16));
                            AxisPolylineID = db.GetObjectId(false, tempHandle, 0);

                            break;

                        case "master":

                            MasterPolylineID = objectId;


                            Number = ListXdata[1].value;                                                   // номер 

                            AxisHandle = new Handle(Convert.ToInt64(ListXdata[2].value, 16)).ToString();   // handle привязяанной оси

                            tempHandle = new Handle(Convert.ToInt64(ListXdata[3].value, 16));              // handle slave полилинии

                            if (!string.Equals(tempHandle.ToString(), "0"))                                // мультилиния
                            {
                                SlavePolylineID = db.GetObjectId(false, tempHandle, 0);                    // Id slave полилинии
                                Type = TsoddElement.Mline;
                            }
                            else
                            {
                                Type = TsoddElement.Line;
                            }

                            tempHandle = new Handle(Convert.ToInt64(ListXdata[4].value, 16));               // handle mtext
                            MtextID = db.GetObjectId(false, tempHandle, 0);                                 // Id mtext

                            tempString = Convert.ToString(ListXdata[5].value);                              // материал
                            if (!string.IsNullOrEmpty(tempString)) Material = tempString;

                            tempString = Convert.ToString(ListXdata[6].value);                              // наличие
                            if (!string.IsNullOrEmpty(tempString))
                            {
                                Existence = tempString;
                                //switch (tempString)
                                //{
                                //    case "Нанести": Existence = "Требуется нанести"; break;
                                //    case "Демаркировать": Existence = "Требуется демаркировать"; break;
                                //    default: Existence = tempString; break;
                                //}
                            }

                            break;

                        case "slave":

                            tempHandle = new Handle(Convert.ToInt64(ListXdata[1].value, 16));               // handle master полилинии
                            MasterPolylineID = db.GetObjectId(false, tempHandle, 0);                        // Id master полилинии

                            // читаем Xdata у мастер полилинии и берем данные уже из master полилинии
                            ListXdata = AutocadXData.ReadXData(MasterPolylineID);
                            // если по какой-то причине XDatra master полилинии пустая, то выходим
                            if (ListXdata.Count == 0) break;

                            tempHandle = new Handle(Convert.ToInt64(ListXdata[3].value, 16));               // handle slave полилинии

                            if (!string.Equals(tempHandle.ToString(), "0"))                                 // мультилиния
                            {
                                SlavePolylineID = db.GetObjectId(false, tempHandle, 0);                     // Id slave полилинии
                                Type = TsoddElement.Mline;
                            }
                            else
                            {
                                Type = TsoddElement.Line;
                            }

                            tempHandle = new Handle(Convert.ToInt64(ListXdata[4].value, 16));               // handle mtext
                            MtextID = db.GetObjectId(false, tempHandle, 0);                                 // Id mtext

                            break;
                    }

                    tr.Commit();
                }
            }

        }
    }





}