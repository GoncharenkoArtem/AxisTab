
using System;
using System.Drawing;
using System.Windows;

using System.Windows.Media.Imaging;
using ACAD = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using System.Collections.Generic;
using System.Windows.Interop;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.EditorInput;
using System.IO;
using System.Reflection;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Windows.Documents;
using System.Windows.Controls;
using System.Security.Cryptography;
using System.Xml.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Interop;
using static System.Net.Mime.MediaTypeNames;
using Autodesk.AutoCAD.Interop.Common;
using ACAD_test;
using System.Security.Policy;
using Autodesk.AutoCAD.GraphicsSystem;
using System.Diagnostics;
using System.Windows.Media.Media3D;
using Autodesk.AutoCAD.DatabaseServices.Filters;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.Windows.Data;
using System.Windows.Media;
using System.Linq;
using Autodesk.Windows;
//using System.Windows.Forms;




namespace TSODD
{
    internal static class TsoddBlock
    {
        private static string dllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        private static string dwgPath = Path.Combine(dllPath, "Support","blocks.dwg");
        public static bool blockInsertFlag = true;


        //// метод проверяет существует ли блок "AXIS_BLOCK", если его нет, то создает
        //private static void CreateBlock()
        //{
        //    var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
        //    var db = doc.Database;
        //    const string blockName = "AXIS_BLOCK"; // имя блока

        //    using (var tr = db.TransactionManager.StartTransaction())
        //    {
        //        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

        //        //ObjectId btrId;
        //        if (bt.Has(blockName)) return;

        //        // Создаём пустое определение блока
        //        bt.UpgradeOpen();
        //        var btr = new BlockTableRecord { Name = blockName };
        //        bt.Add(btr);
        //        tr.AddNewlyCreatedDBObject(btr, true);

        //        // Добавляем невидимые атрибуты 
        //        AddHiddenAttr(btr, tr, tag: "POLYHANDLE", prompt: "Handle полилинии оси", defaultValue: "", 0);
        //        AddHiddenAttr(btr, tr, tag: "AXISNAME", prompt: "Наименование оси", defaultValue: "", 1);
        //        //AddHiddenAttr(btr, tr, tag: "STARTPOINT", prompt: "Начальная точка", defaultValue: "", 2);
        //        AddHiddenAttr(btr, tr, tag: "REVERSE", prompt: "Реверсивное направление", defaultValue: "", 3);
        //        AddHiddenAttr(btr, tr, tag: "PK", prompt: "Начальный пикет", defaultValue: "", 4);

        //        tr.Commit();
        //    }
        //}


        // метод создания нового атрибута
        private static void AddHiddenAttr(BlockTableRecord btr, Transaction tr, string tag, string prompt, string defaultValue, int numberOfAttrr, double height = 2.5)
        {
            var ad = new AttributeDefinition
            {
                Position = new Point3d(0, -numberOfAttrr * 2.5, 0),
                Tag = tag.ToUpperInvariant(),
                Prompt = prompt ?? string.Empty,
                TextString = defaultValue ?? string.Empty,
                Height = height,
                Invisible = true,
                Visible = false,
                Constant = false,
                LockPositionInBlock = true
            };

            btr.AppendEntity(ad);
            tr.AddNewlyCreatedDBObject(ad, true);
        }


        //// метод проверяет, существует ли определение блока по имени блока, и создает либо обновляет его (атрибуты)
        //public static void AddAxisBlock(Axis axis)
        //{
        //    var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
        //    var db = doc.Database;
        //    const string blockName = "AXIS_BLOCK"; // имя блока

        //    using (ACAD.DocumentLock docLock = doc.LockDocument()) // Блокируем документ
        //    {
        //        // проверяем есть ли блок в реестре, если нет, то создаем его
        //        CreateBlock();

        //        // Ищем вхождение блока с нужными атрибутами (по наименованию оси)
        //        using (var tr = db.TransactionManager.StartTransaction())
        //        {
        //            // ищем Id блока
        //            ObjectId btrId;
        //            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

        //            btrId = bt[blockName];

        //            // определение блока
        //            var def = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

        //            // ищем нужный блок
        //            var axisBlockRef = FindAxisBlock(def, tr, axis.PolyHandle);

        //            if (axisBlockRef == null) // если не нашли, то вставляем новый блок в 0 
        //            {
        //                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
        //                axisBlockRef = new BlockReference(Point3d.Origin, btrId);
        //                ms.AppendEntity(axisBlockRef);
        //                tr.AddNewlyCreatedDBObject(axisBlockRef, true);

        //                SetAttrValues(axisBlockRef, tr);
        //            }

        //            // заполняем/обновляем атрибуты блока
        //            ChangeAttribute(tr, axisBlockRef, "POLYHANDLE", axis.PolyHandle.ToString());
        //            ChangeAttribute(tr, axisBlockRef, "AXISNAME", axis.Name);
        //            //ChangeAttribute(tr, axisBlockRef, "STARTPOINT", $"{axis.StartPoint.X},{axis.StartPoint.Y},{axis.StartPoint.Z}");
        //            ChangeAttribute(tr, axisBlockRef, "REVERSE", axis.ReverseDirection.ToString());
        //            ChangeAttribute(tr, axisBlockRef, "PK", axis.StartPK.ToString());

        //            tr.Commit();

        //        }
        //    }
        //}


        //// метод поиска нужного блока по Handle
        //private static BlockReference FindAxisBlock(BlockTableRecord def, Transaction tr, Handle axishandle)
        //{
        //    foreach (ObjectId brId in def.GetBlockReferenceIds(/*directOnly*/ true, /*includeErased*/ false))
        //    {
        //        // получаем экземпляр блока
        //        var br = (BlockReference)tr.GetObject(brId, OpenMode.ForRead);

        //        // открываем определение, на которое указывает вставка, для проверки Xref
        //        var brDef = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);

        //        // если это внешняя ссылка (attach/overlay) — пропускаем
        //        if (brDef.IsFromExternalReference || brDef.IsFromOverlayReference)
        //            continue;

        //        // Обход атрибутов вставки 
        //        foreach (ObjectId arId in br.AttributeCollection)
        //        {
        //            var ar = (AttributeReference)tr.GetObject(arId, OpenMode.ForRead);

        //            if (ar.Tag.Equals("POLYHANDLE", StringComparison.OrdinalIgnoreCase) && ar.TextString == axishandle.ToString())
        //            { return br; }

        //        }
        //    }
        //    return null;
        //}


        // метод изменения отрибутов блока
        private static void ChangeAttribute(Transaction tr, BlockReference br, string tag, string value)
        {
            // Обход атрибутов вставки 
            foreach (ObjectId arId in br.AttributeCollection)
            {
                var ar = (AttributeReference)tr.GetObject(arId, OpenMode.ForWrite);
                if (ar.Tag.Equals(tag, StringComparison.OrdinalIgnoreCase)) ar.TextString = value;
            }
        }

        // метод добавления атрибутов в новый блок
        private static void SetAttrValues(BlockReference br, Transaction tr, Dictionary<string, string> tagList = null)
        {

            // Текущий масштаб аннотаций
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var occ = db.ObjectContextManager.GetContextCollection("ACDB_ANNOTATIONSCALES");
            var cur = occ.CurrentContext as AnnotationScale ?? db.Cannoscale as AnnotationScale;


            var btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
            foreach (ObjectId id in btr)
            {
                if (tr.GetObject(id, OpenMode.ForRead) is AttributeDefinition ad && !ad.Constant)
                {
                    var ar = new AttributeReference();
                    ar.SetDatabaseDefaults();
                    ar.SetAttributeFromBlock(ad, br.BlockTransform);

                    // если надо заполнить атрибуты
                    if (tagList != null)
                    {
                        // если есть подходящий тег
                        if (tagList.TryGetValue(ar.Tag, out string val))
                        {
                            ar.TextString = val;
                            tagList.Remove(ar.Tag);
                        }
                    }

                    // добавляем аннатотивность в атрибут
                    if (ad.Annotative == AnnotativeStates.True || btr.Annotative == AnnotativeStates.True || br.Annotative == AnnotativeStates.True)
                    {
                        if (!ar.Invisible) ar.AddContext(cur);
                    }

                    br.AttributeCollection.AppendAttribute(ar);
                    tr.AddNewlyCreatedDBObject(ar, true);

                    // выравнивание после присвоения текста
                    ar.AdjustAlignment(db);
                }
            }

            // проверяем все ли атрибуты заполнелись
            if (tagList != null && tagList.Count != 0)
            {
                string att = string.Empty;
                foreach (var a in tagList) att += $"\n тег: {a.Key} , значение: {a.Value} \n";
                var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage("\n Не получилось заполнить следующие атрибуты блока: \n " + att);
                return;
            }
        }


        //// Удалить вхождения конкретного блока по совпадению значения атрибута (TAG = value)
        //public static void DeleteAxisBlock(Axis axis)
        //{
        //    var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
        //    var db = doc.Database;
        //    const string blockName = "AXIS_BLOCK"; // имя блока

        //    using (doc.LockDocument())
        //    using (var tr = db.TransactionManager.StartTransaction())
        //    {
        //        // Ищем вхождение блока с нужными атрибутами (по наименованию оси)
        //        ObjectId btrId;
        //        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

        //        btrId = bt[blockName];

        //        // определение блока
        //        var def = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

        //        // ищем нужный блок
        //        var axisBlockRef = FindAxisBlock(def, tr, axis.PolyHandle);

        //        if (axisBlockRef == null) return;

        //        try
        //        {
        //            axisBlockRef.UpgradeOpen();
        //            if (!axisBlockRef.IsErased) axisBlockRef.Erase();

        //        }
        //        catch (Autodesk.AutoCAD.Runtime.Exception)
        //        {
        //            // при необходимости — лог ex.ErrorStatus
        //        }
        //        tr.Commit();
        //    }

        //}


























        // метод импортирования блоков с шаблона с настройками атрибутов
        private static void InsertBlockByNameWithTags(Point3d insertPoint, string name, Dictionary<string, string> tags)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (var blockDb = new Database(false, true))     // новая пустая база
            {
                blockDb.ReadDwgFile(dwgPath, FileShare.Read, true, "");     // считываем в базу файл по указанному пути

                using (var trExt = blockDb.TransactionManager.StartTransaction())
                {
                    var btExt = (BlockTable)trExt.GetObject(blockDb.BlockTableId, OpenMode.ForRead);

                    if (!btExt.Has(name))
                    {
                        ed.WriteMessage($"\n Блок \"{name}\" не найден.");
                        blockInsertFlag = true;
                        return;

                    }

                    var extBtrId = btExt[name];

                    // Импортируем определение блока в текущий чертёж
                    using (doc.LockDocument())
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        // Скопировать определение и все его зависимости (слои, стили и т.п.)
                        var ids = new ObjectIdCollection { extBtrId };
                        var mapping = new IdMapping();

                        // ownerId — это BlockTable текущей БД
                        blockDb.WblockCloneObjects(
                        ids,
                        db.BlockTableId,
                        mapping,
                        DuplicateRecordCloning.Replace, // Replace=перезаписать если есть; Ignore=оставить существующий
                        false
                        );

                        // Получаем ObjectId склонированного (или существующего) определения в текущей БД
                        var pair = mapping[extBtrId];
                        var destBtrId = pair.Value;

                        blockInsertFlag = false;

                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        var br = new BlockReference(insertPoint, destBtrId);

                        ms.AppendEntity(br);
                        tr.AddNewlyCreatedDBObject(br, true);



                        var ocm = db.ObjectContextManager;
                        var occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");

                        // Текущий масштаб CANNOSCALE
                        var curSC = occ.CurrentContext as AnnotationScale;
                        if (curSC != null && !br.HasContext(curSC))

                            br.AddContext(curSC);   // добавляем масштаб в блок

                        // заполняем атрибуты
                        SetAttrValues(br, tr, tags);

                        br.RecordGraphicsModified(true);

                        blockInsertFlag = true;  // можно опять вставлять блоки

                        tr.Commit();
                    }
                    trExt.Commit();
                }
            }
        }

        // метод вставки блока стойки
        public static void InsertStandBlock(string name)
        {

            if (blockInsertFlag == false) return;
            blockInsertFlag = false;  // переменная для отслеживанияч двойного нажатия кнопки

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // проверка на наличие осей в файле
            if (TsoddHost.Current.axis.Count == 0)
            {
                ed.WriteMessage("\n В чертеже не найдено осей. Для построения новой оси введите команду \"NEWAXIS\" или нажмите кнопку на панели \"Новая ось\" \n");
                blockInsertFlag = true;
                return;
            }

            // если не выбрана текущая ось, то и не будем ничего вставлять
            if (TsoddHost.Current.currentAxis == null)
            {
                ed.WriteMessage("\n Ошибка. не выбрана текущая ось. \n");
                blockInsertFlag = true;
                return;
            }

            // Запрашиваем точку и вставляем BlockReference
            var ppo = ed.GetPoint("\n Укажите точку вставки блока (Esc - выход): ");
            if (ppo.Status != PromptStatus.OK)
            {
                blockInsertFlag = true;
                return;
            }

            // словарь для заполнения тегов блока
            var tagList = new Dictionary<string, string> { ["ОСЬ"] = TsoddHost.Current.currentAxis.Name };

            // вставляем блок стойки
            InsertBlockByNameWithTags(ppo.Value, name, tagList);

        }


        // метод вставки блока знака
        public static void InsertSignBlock(string name)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // соберем все блоки стоек в отдельный словарь где будут точки вставки и ID блока
            Dictionary<ObjectId, Point3d> standBlocks = new Dictionary<ObjectId, Point3d>();

            // последний подсвеченный блок
            ObjectId lastHighlightedId = ObjectId.Null;

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // таблица со всеми блоками в чертеже
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                foreach (var blockID in bt)
                {
                    var def = (BlockTableRecord)tr.GetObject(blockID, OpenMode.ForRead); // определение блока

                    if (def.IsFromExternalReference || def.IsFromOverlayReference) continue; // игнорируем внешние ссылки

                    bool match = false; // переменная, чтобы понимать, что мы нашли нужный тип блока 

                    foreach (ObjectId id in def)
                    {
                        if (tr.GetObject(id, OpenMode.ForRead) is AttributeDefinition ad && !ad.Constant)
                        {
                            if (ad.Tag.Equals("STAND", StringComparison.OrdinalIgnoreCase))
                            { match = true; break; }    // есть совпадение, перрываем перебор
                        }
                    }

                    if (match == false) continue; // если не нашли подходящий тег, то продолжаем поиск 

                    // перебираем все вхождения блока 
                    foreach (ObjectId brId in def.GetBlockReferenceIds(/*directOnly*/ true, /*includeErased*/ false))
                    {
                        var bref = (BlockReference)tr.GetObject(brId, OpenMode.ForRead);
                        standBlocks.Add(brId, bref.Position);
                    }
                }
                tr.Commit();
            }

            // подписывемся на курсор
            ed.PointMonitor += Ed_PointMonitor;

            // создаем промпт выбора точки вставки блока
            var ppr = ed.GetPoint("\n Укажите точку вставки блока (Esc - выход): ");

            // обработчик события курсора
            void Ed_PointMonitor(object sender, PointMonitorEventArgs e)
            {

                Point3d ucsPt = e.Context.ComputedPoint;
                Point3d cursorWcs = ucsPt.TransformBy(ed.CurrentUserCoordinateSystem.Inverse());  // текущее положение курсора

                double maxRadius = 5000;    // максимальный радиус поиска
                if (db.Insunits == UnitsValue.Meters) maxRadius = 5; // если в чертеже размерность в метрах, то меняем радиус поиска стойки

                ObjectId matchBlockID = ObjectId.Null;
                double distance = double.MaxValue;

                foreach (var block in standBlocks)
                {
                    double curDist = block.Value.DistanceTo(cursorWcs);
                    if (curDist <= maxRadius && curDist < distance)
                    {
                        distance = curDist;
                        matchBlockID = block.Key;
                    }
                }

                // подсветка
                SetHighlight(matchBlockID);

            }

            // подсветка вкл/выкл 
            void SetHighlight(ObjectId newId)
            {
                if (newId == lastHighlightedId) return;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // снять со старого
                    if (!lastHighlightedId.IsNull)
                    {
                        if (tr.GetObject(lastHighlightedId, OpenMode.ForRead, false) is Entity oldEnt)
                            oldEnt.Unhighlight();
                    }

                    // включить на новом
                    if (!newId.IsNull)
                    {
                        if (tr.GetObject(newId, OpenMode.ForRead, false) is Entity newEnt)
                            newEnt.Highlight();
                    }

                    lastHighlightedId = newId;
                    tr.Commit();
                }
            }

            // бесконечный цикл для обработки промпта
            while (true)
            {
                try
                {
                    if (ppr.Status != PromptStatus.OK) return;  // если пользователь отменил промпт с точкой, то выходим

                    // последний подсвеченый блок это lastHighlightedBlock. Если он не null, то можно брать данные с него
                    if (!lastHighlightedId.IsNull)
                    {
                        using (doc.LockDocument())
                        {
                            using (var tr = db.TransactionManager.StartTransaction())
                            {

                                var bref = (BlockReference)tr.GetObject(lastHighlightedId, OpenMode.ForWrite);

                                var tagList = new Dictionary<string, string> { ["STANDHANDLE"] = bref.Handle.ToString() };
                                InsertBlockByNameWithTags(ppr.Value, name, tagList);

                                tr.Commit();
                            }
                        }
                        return; // успех
                    }
                    else  // нет подходящего блока, тогда вызываем новый промпт для выбора блока стойки
                    {
                        while (true)
                        {
                            var ppr2 = ed.GetPoint("\n Укажите блок стойки к которому нужно привязать знак (Esc - выход): ");
                            if (ppr2.Status != PromptStatus.OK)
                                return;

                            if (!lastHighlightedId.IsNull)
                            {
                                using (var tr = db.TransactionManager.StartTransaction())
                                {
                                    var bref = (BlockReference)tr.GetObject(lastHighlightedId, OpenMode.ForRead);
                                    var tagList = new Dictionary<string, string> { ["STANDHANDLE"] = bref.Handle.ToString() };

                                    // вставляем блок
                                    InsertBlockByNameWithTags(ppr.Value, name, tagList);
                                    tr.Commit();

                                }
                                return; // успех
                            }
                        }
                    }
                }
                finally
                {
                    // отписываемся от курсора
                    ed.PointMonitor -= Ed_PointMonitor;
                    // выключаем подсветку
                    SetHighlight(ObjectId.Null);
                }
            }
        }


        //  **************************************************************   ADD / DELETE  BLOCKS   **************************************************************    //

        // метод загрузки блока в базу
        public static string AddBlockToBD(string templateName)
        {
            string blockName = null;

            string intBlockName = string.Empty;
            ObjectIdCollection idsToClone = null;

            double entityWidth = 0;     // переменная для отслеживания ширины всех примитивов блока
            double entityHeight = 0;    // переменная для отслеживания высоты всех примитивов блока
            double minX = 0;            // переменная для ослеживания начального X примитивов блока
            double minY = 0;            // переменная для ослеживания начального Y примитивов блока


            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // промпт выбора исходного блока 
            PromptEntityOptions peo = new PromptEntityOptions("\n Выберите блок для экспорта в БД (Esc - выход):");
            peo.SetRejectMessage(" \n Выбранный элемент должен быть блоком");
            peo.AddAllowedClass(typeof(BlockReference), exactMatch: false);
            var res = ed.GetEntity(peo);

            if (res.Status != PromptStatus.OK) return null; // если пользователь отменил промпт то выходим

            ObjectId intBtrId = res.ObjectId;

            Bitmap blockbmp;

            try
            {
                // берем данные с исходного блока
                using (doc.LockDocument())
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        // получаем экземпляр блока
                        var intBr = (BlockReference)tr.GetObject(intBtrId, OpenMode.ForRead);

                        // если блок динамический — берем базовое определение
                        ObjectId defId = intBr.IsDynamicBlock
                            ? intBr.DynamicBlockTableRecord
                            : intBr.BlockTableRecord;

                        // открываем определение, на которое указывает вставка, для проверки Xref
                        var intBtr = (BlockTableRecord)tr.GetObject(defId, OpenMode.ForRead);

                        // имя исходного блока
                        intBlockName = intBtr.Name;

                        // Id исходного блока
                        intBtrId = intBtr.ObjectId;

                        // копируем сущности из исходного блока 
                        idsToClone = new ObjectIdCollection();
                        foreach (ObjectId id in intBtr)
                        {
                            var dbo = tr.GetObject(id, OpenMode.ForRead);
                            if (dbo is AttributeDefinition) continue;           // кроме атрибутов
                            if (dbo is DBText || dbo is MText) continue;        // кроме текста и Мтекста
                            if (dbo.GetType().FullName == "Autodesk.AutoCAD.DatabaseServices.MTextAttributeDefinition") continue;

                            if (dbo is Entity) idsToClone.Add(id);
                        }

                        // временный блок для картинки
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                        var tempBtr = new BlockTableRecord();
                        tempBtr.Name = "tempBlockForPreviewIcon"; // анонимное имя, чтобы не конфликтовать

                        ObjectId tempBtrId = bt.Add(tempBtr);
                        tr.AddNewlyCreatedDBObject(tempBtr, true);

                        var tempMap = new IdMapping();
                        db.DeepCloneObjects(idsToClone, tempBtrId, tempMap, false);

                        // генерим картинку bmp для previewIcon заодно понимаем положение примитивов
                        blockbmp = GetBlockBitmap(tempBtrId, out entityWidth, out entityHeight, out minX, out minY, 32);

                        tempBtr.UpgradeOpen();
                        tempBtr.Erase(true);

                    }
                }

                // копируем шаблон из файла с блоками, шаманим с ним и сохраняем под новым именем
                using (var extDb = new Database(false, true))
                {

                    // Открываем внешний DWG для записи
                    extDb.ReadDwgFile(dwgPath, FileShare.None, false, "");
                    extDb.CloseInput(true);

                    // Находим шаблон в внешней БД
                    ObjectId tplId;
                    using (var trExt = extDb.TransactionManager.StartTransaction())
                    {
                        // проверка наличия шаблона
                        var btExt = (BlockTable)trExt.GetObject(extDb.BlockTableId, OpenMode.ForRead);
                        if (!btExt.Has(templateName))
                        {
                            ed.WriteMessage($"\nВ БД не найден шаблон \"{templateName}\".");
                            return null;
                        }

                        tplId = btExt[templateName];

                        // проверка имени добавляемого блока в БД
                        if (btExt.Has(intBlockName))
                        {
                            // спрашиваем заменить ли определение блока или выход
                            PromptKeywordOptions pso = new PromptKeywordOptions("\n Такой блок уже есть в БД. Переопределить его?: ");
                            pso.Keywords.Add("Да");
                            pso.Keywords.Add("Нет");
                            pso.AllowArbitraryInput = false;

                            PromptResult pres = ed.GetKeywords(pso);

                            if (pres.StringResult != "Да")  // выходим, если не хотим переопределить блок
                            {
                                trExt.Abort();
                                return null;
                            }
                        }
                        trExt.Commit();
                    }

                    // Клонируем шаблон во временную БД
                    using (var tmpDb = new Database(true, true))
                    {
                        var tempMap = new IdMapping();
                        extDb.WblockCloneObjects(
                            new ObjectIdCollection(new[] { tplId }),
                            tmpDb.BlockTableId,
                            tempMap,
                            DuplicateRecordCloning.MangleName, // создаст копию записи блока
                            false
                        );

                        // Получаем id клона в temp
                        ObjectId tmpBtrId = ObjectId.Null;
                        foreach (IdPair p in tempMap)
                            if (p.Key == tplId) { tmpBtrId = p.Value; break; }


                        // Переименовываем клон в temp в целевое имя
                        using (var trTmp = tmpDb.TransactionManager.StartTransaction())
                        {
                            var btTmp = (BlockTable)trTmp.GetObject(tmpDb.BlockTableId, OpenMode.ForWrite);
                            var tmpBtr = (BlockTableRecord)trTmp.GetObject(tmpBtrId, OpenMode.ForWrite);
                            tmpBtr.Name = intBlockName;
                            trTmp.Commit();
                        }

                        // Клонируем блок из временной БД во внешнюю БД
                        var mapBack = new IdMapping();
                        tmpDb.WblockCloneObjects(
                            new ObjectIdCollection(new[] { tmpBtrId }), // источник: tempBD
                            extDb.BlockTableId,                         // владелец: BlockTable внешней БД
                            mapBack,
                            DuplicateRecordCloning.Replace,
                            false
                        );

                        // Получаем созданный BTR во внешней БД
                        ObjectId newBtrId = ((IdPair)mapBack[tmpBtrId]).Value;

                        // Копируем сущности
                        var map = new IdMapping();
                        db.WblockCloneObjects(
                            idsToClone,            // объекты из текущего чертежа
                            newBtrId,              // владелец: новый BlockTableRecord во внешней БД
                            map,
                            DuplicateRecordCloning.Ignore, // слои/типы/блоки-дочерние подтянутся
                            false
                        );

                        // делаем подмену картинки previewIcon
                        using (var trExt = extDb.TransactionManager.StartTransaction())
                        {
                            var newBtr = (BlockTableRecord)trExt.GetObject(newBtrId, OpenMode.ForWrite);
                            newBtr.PreviewIcon = blockbmp;  // меняем previewIcon у блока
                            newBtr.Annotative = AnnotativeStates.True;

                            // настраиваем блок в зависимости от его типа
                            switch (templateName)
                            {
                                case "STAND_TAMPLATE":

                                    break;

                                case "SIGN_TEMPLATE":

                                    // спрашиваем сдвоенный ли знак
                                    bool doubled = false;

                                    PromptKeywordOptions pso = new PromptKeywordOptions("\n Блок состоит из одного занака?: ");
                                    pso.Keywords.Add("Да");
                                    pso.Keywords.Add("Нет");
                                    pso.Keywords.Add("Выход");

                                    pso.AllowArbitraryInput = false;
                                    PromptResult pres = ed.GetKeywords(pso);

                                    if (pres.StringResult == "Выход")
                                    {
                                        return null;
                                    }

                                    if (pres.StringResult != "Да")  // выходим, если не хотим переопределить блок
                                    {
                                        doubled = true;
                                    }

                                    // меняем атрибут GROUP для знака
                                    foreach (ObjectId id in newBtr)
                                    {
                                        if (trExt.GetObject(id, OpenMode.ForWrite) is AttributeDefinition ad && !ad.Constant)
                                        {

                                            if (ad.Tag.Equals("НОМЕР_ЗНАКА", StringComparison.OrdinalIgnoreCase))
                                            {
                                                double offset = Math.Max(entityWidth, entityHeight) / 20;
                                                ad.TextString = intBlockName;
                                                ad.Position = new Point3d(minX + entityWidth + offset, minY + entityHeight / 2 - ad.Height / 2, 0);
                                            }

                                            if (ad.Tag.Equals("GROUP", StringComparison.OrdinalIgnoreCase))
                                            {
                                                ad.TextString = TsoddHost.Current.currentSignGroup; // привязываем к текущей группе знаков
                                            }

                                            if (ad.Tag.Equals("DOUBLED", StringComparison.OrdinalIgnoreCase))
                                            {
                                                ad.TextString = doubled.ToString();
                                            }
                                        }
                                    }

                                    break;
                            }
                            blockName = newBtr.Name;
                            trExt.Commit();
                        }
                        // Сохраняем внешний файл
                        extDb.SaveAs(dwgPath, DwgVersion.Current);

                    }
                }
            }

            catch
            {
                ed.WriteMessage($"\n Блок \"{intBlockName}\" не получилось добавить в базу. \n");
                return null;
            }
            ed.WriteMessage($"\n Блок \"{intBlockName}\" успешно добавлен в базу. \n");
            return blockName;
        }


        public static void DeleteBlockFromBD(string blockName)
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (var extDb = new Database(false, true))
            {
                // Открываем внешний DWG для записи
                extDb.ReadDwgFile(dwgPath, FileShare.None, false, "");
                extDb.CloseInput(true);

                using (var trExt = extDb.TransactionManager.StartTransaction())
                {
                    // проверка наличия шаблона
                    var btExt = (BlockTable)trExt.GetObject(extDb.BlockTableId, OpenMode.ForRead);
                    if (!btExt.Has(blockName))
                    {
                        ed.WriteMessage($"\n В БД не найден блок \"{blockName}\".");
                        return;
                    }
                    else
                    {
                        PromptKeywordOptions pso = new PromptKeywordOptions($"\n Вы точно хотите удалить блок \"{blockName}\" из БД?: ");
                        pso.Keywords.Add("Да");
                        pso.Keywords.Add("Нет");
                        pso.AllowArbitraryInput = false;

                        PromptResult pres = ed.GetKeywords(pso);

                        if (pres.StringResult != "Да")  // выходим, если не хотим удалять блок
                        {
                            trExt.Abort();
                            return;
                        }

                        // получаем ID определения блока 
                        var btrID = btExt[blockName];

                        var btr = (BlockTableRecord)trExt.GetObject(btrID, OpenMode.ForWrite);
                        if (!btr.IsErased) btr.Erase();
                    }

                    trExt.Commit();
                }

                // Сохраняем внешний файл
                extDb.SaveAs(dwgPath, DwgVersion.Current);
            }
        }



        //  *****************************************  Работа с Bitmap ******************************************

        //подключаем WinAPI-функцию DeleteObject из gdi32.dll.
        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        // конвертирует из Bitmap в BitmapSource для Ribbon
        public static BitmapSource ToImageSource(Bitmap bmp)
        {
            if (bmp == null) return null;
            IntPtr h = bmp.GetHbitmap();        // получаем дескриптор HBITMAP
            try
            {
                return Imaging.CreateBitmapSourceFromHBitmap(    //(WPF) строит BitmapSource из HBITMAP
                    h, IntPtr.Zero, Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
            }
            finally
            {
                DeleteObject(h);    // эта функция освобождает GDI-ресурсы (в нашем случае — HBITMAP), чтобы не было утечек.
            }
        }

        // метод делает snapShot (чнимок настроенного вида блока) и возвращает bmp.
        private static Bitmap GetBlockBitmap(ObjectId btrId, out double entityWidth, out double entityHeight, out double minX, out double minY, byte px = 32)
        {

            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            Bitmap bmp = null;
            Manager manager = ACAD.Application.DocumentManager.MdiActiveDocument.GraphicsManager;

            KernelDescriptor descriptor = new KernelDescriptor();

            descriptor.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"));
            GraphicsKernel m_pGraphicsKernel = Manager.AcquireGraphicsKernel(descriptor);
            Device device = manager.CreateAutoCADOffScreenDevice(m_pGraphicsKernel);
            Model model = manager.CreateAutoCADModel(m_pGraphicsKernel);

            using (device)
            {
                device.OnSize(new System.Drawing.Size(px, px));
                using (model)
                {

                    View view = new Autodesk.AutoCAD.GraphicsSystem.View();
                    device.Add(view);

                    Point3d entityPosition;

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        // определение блока 
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                        // создаем вхождение блока 
                        var br = new BlockReference(Point3d.Origin, btrId);

                        view.Add(br, model);    // добавляем на вид блок

                        // определим габариты вхождения блока
                        var bounds = br.GeometricExtents;

                        minX = bounds.MinPoint.X;
                        minY = bounds.MinPoint.Y;
                        entityWidth = Math.Round(bounds.MaxPoint.X - minX,2);
                        entityHeight = Math.Round(bounds.MaxPoint.Y - minY,2);
                        entityPosition = new Point3d((bounds.MaxPoint.X + bounds.MinPoint.X) * 0.5, (bounds.MaxPoint.Y + bounds.MinPoint.Y) * 0.5, 0);

                        try
                        {
                            view.SetView(entityPosition, entityPosition, view.UpVector, entityWidth, entityHeight);
                            view.ZoomExtents(bounds.MinPoint, bounds.MaxPoint);
                        }
                        catch (System.Exception ex) { Debug.Print(ex.Message); }
                        ;
                        
                        System.Drawing.Rectangle rect = new System.Drawing.Rectangle(0, 0, px, px);
                        bmp = view.GetSnapshot(rect);
                    }
                }
            }
            return bmp;
        }


        // метод получает список имен блоков и их previewIcon, которые соответсвуют тегу (для заполнения RibbonSplitButton)
        public static List<(string name, BitmapSource img)> GetListOfBlocks(string typeTag, string group)
        {
            List<(string name, BitmapSource img)> result = new List<(string, BitmapSource)>(); // возвращаемый список блоков (имя и картинка привью)

            using (var blockDb = new Database(false, true))     // новая пустая база
            {
                blockDb.ReadDwgFile(dwgPath, FileShare.Read, true, "");     // считываем в базу файл по указанному пути

                using (var tr = blockDb.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(blockDb.BlockTableId, OpenMode.ForRead);

                    foreach (var blockID in bt)
                    {
                        var def = (BlockTableRecord)tr.GetObject(blockID, OpenMode.ForRead);

                        if (def.Name.Contains("_TEMPLATE")) continue;

                        bool match = false;

                        foreach (var id in def)
                        {
                            if (tr.GetObject(id, OpenMode.ForRead) is AttributeDefinition at)
                            {
                                if (at.Tag.Equals(typeTag, StringComparison.OrdinalIgnoreCase) && group == null)     // нашли нужный блок с соотвествующим тегом
                                {
                                    match = true;
                                    break;
                                }

                                if (at.Tag.Equals("GROUP", StringComparison.OrdinalIgnoreCase) && at.TextString == group)     // нашли нужный блок с соотвествующим тегом
                                {
                                    match = true;
                                    break;
                                }
                            }
                        }

                        // если блок подходит
                        if (match)
                        {
                            var img = TsoddBlock.ToImageSource(def.PreviewIcon);  // генерим bitmasource для привью блока
                            result.Add((def.Name, img));       // записываем данные блока в список
                        }
                    }
                }
            }

            return result;
        }

        //  **************************************************************   РАБОТА СО СТОЙКАМИ   **************************************************************    //

        // перепривязывает блоки стоек к выбранной оси
        public static void ReBindStandBlockToAxis()
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var objectIdList = RibbonInitializer.Instance.GetAutoCadSelectionObjectsId(new List<string> { "INSERT" });

            // постфильтр, проверяем есть ли в блоках тег "STAND", порлучем список с Id блоков
            List<ObjectId> standBlocks = GetBloclListIdByTagFromSelection(db, objectIdList, "STAND");

            if (standBlocks == null || standBlocks.Count == 0) // если не нашли блоки стоек, то выходим
            {
                ed.WriteMessage("\n Не выбрано ни одного блока стоек \n");
                return;
            }
            else
            {
                ed.WriteMessage($"\n Найдено {standBlocks.Count} блоков стоек \n");
            }

            var selectedAxis = TsoddCommands.SelectAxis();  // готовая команда вернет выбранную ось
            if (selectedAxis == null) return;


            // перебором меняем нужный тег привязки к оси
            using (var lok = doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var blockId in standBlocks)
                {
                    var bref = (BlockReference)tr.GetObject(blockId, OpenMode.ForRead);
                    ChangeAttribute(tr, bref, "ОСЬ", selectedAxis.Name);
                }

                tr.Commit();
            }

            ed.WriteMessage($"\n Блоки перепривязаны к оси {selectedAxis.Name} \n");
        }



        // получает списиок Id блоков из выделения рамкой, которые соответствуют указанному тегу
        private static List<ObjectId> GetBloclListIdByTagFromSelection(Database db, List<ObjectId> selection, string tag)
        {
            if (selection == null) return null;

            List<ObjectId> result = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var id in selection)
                {
                    var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    if (br == null) continue;

                    foreach (ObjectId arId in br.AttributeCollection)
                    {
                        var ar = tr.GetObject(arId, OpenMode.ForRead) as AttributeReference;
                        if (ar == null) continue;

                        if (string.Equals(ar.Tag, tag, StringComparison.OrdinalIgnoreCase)) result.Add(id);
                    }
                }
                tr.Commit();
            }
            return result;
        }


        //  **************************************************************   РАБОТА СО ЗНАКАМИ   **************************************************************    //


        // Заполнение группы знаков
        public static void FillSignsGroups()
        {
            using (var extDb = new Database(false, true))
            {
                // Открываем внешний DWG для записи
                extDb.ReadDwgFile(dwgPath, FileShare.None, false, "");
                extDb.CloseInput(true);

                using (var trExt = extDb.TransactionManager.StartTransaction())
                {
                    var groupList = GetListOfSignGroups(trExt, extDb, out BlockTableRecord btr);
                    RibbonInitializer.Instance.signsGroups.Items.Clear();
                    foreach (var group in groupList) RibbonInitializer.Instance.signsGroups.Items.Add(new RibbonButton { Text = group, ShowText = true });
                }
            }
        }


        // добавляет группу знаков в шаблон
        public static void AddSignGroupToBD()
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            PromptStringOptions pso = new PromptStringOptions("\n Введите наименование новой группы знаков (Esc - выход): ");
            pso.AllowSpaces = true;

            var per = ed.GetString(pso);
            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n Ошибка ввода наименования группы...");
                return;
            }

            // новое наименование группы 
            var groupName = per.StringResult;

            using (var extDb = new Database(false, true))
            {
                // Открываем внешний DWG для записи
                extDb.ReadDwgFile(dwgPath, FileShare.None, false, "");
                extDb.CloseInput(true);

                using (var trExt = extDb.TransactionManager.StartTransaction())
                {
                    // список групп в блоке знаков

                    var groupList = GetListOfSignGroups(trExt, extDb, out BlockTableRecord btr);

                    if (btr == null)    // если по какой-то причине у нас нет определения блока
                    {
                        ed.WriteMessage($" \n Ошибка. Не найден блок \"SIGNS_GROUPS\" ");
                        return;
                    }

                    // проверяем нет ли дубликатов
                    var dublicate = groupList.Any(g => string.Equals(g, groupName, StringComparison.OrdinalIgnoreCase));

                    if (dublicate)
                    {
                        ed.WriteMessage($" \n Группа с именем \"{groupName}\" уже существует");
                        return;
                    }

                    MText mText = new MText();
                    mText.Contents = groupName;

                    var id = btr.AppendEntity(mText);
                    trExt.AddNewlyCreatedDBObject(mText, true);

                    trExt.Commit();
                }

                // Сохраняем внешний файл,
                extDb.SaveAs(dwgPath, DwgVersion.Current);
            }

            FillSignsGroups();
        }


        // удаляет группу знаков из шаблона
        public static void DeleteSignGroupFromBD(string groupName)
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            PromptKeywordOptions pso = new PromptKeywordOptions($"\n Вы точно хотите удалить группу знаков \"{groupName}\" и входящие в нее блоки знаков из БД? [Да/Нет]: ");
            pso.Keywords.Add("Да");
            pso.Keywords.Add("Нет");
            pso.AllowArbitraryInput = false;

            PromptResult pres = ed.GetKeywords(pso);

            if (pres.StringResult != "Да")  // выходим, если не хотим удалять группу
            {
                ed.WriteMessage($" \n Ошибка удаления группы знаков ...");
                return;
            }

            // Если таки удаляем группу и знаки
            using (var extDb = new Database(false, true))
            {
                // Открываем внешний DWG для записи
                extDb.ReadDwgFile(dwgPath, FileShare.None, false, "");
                extDb.CloseInput(true);

                using (var trExt = extDb.TransactionManager.StartTransaction())
                {
                    // список групп в блоке знаков
                    var groupList = GetListOfSignGroups(trExt, extDb, out BlockTableRecord btr);

                    if (btr == null)    // если по какой-то причине у нас нет определения блока
                    {
                        ed.WriteMessage($" \n Ошибка. Не найден блок \"SIGNS_GROUPS\" ");
                        return;
                    }

                    // удаляем группу 
                    foreach (var oid in btr)
                    {
                        var obj = trExt.GetObject(oid, OpenMode.ForWrite);
                        if (obj is MText mt)
                        {
                            if (mt.Contents == groupName)
                            {
                                if (!mt.IsErased) mt.Erase();
                            }
                        }
                    }

                    // удаляем все блоки, которые входили в эту группу
                    var bt = (BlockTable)trExt.GetObject(extDb.BlockTableId, OpenMode.ForRead);

                    foreach (var blockID in bt)
                    {
                        var def = (BlockTableRecord)trExt.GetObject(blockID, OpenMode.ForWrite);

                        if (def.Name.Contains("_TEMPLATE")) continue;

                        foreach (var id in def)
                        {
                            if (trExt.GetObject(id, OpenMode.ForRead) is AttributeDefinition at)
                            {
                                if (at.Tag.Equals("GROUP", StringComparison.OrdinalIgnoreCase) && at.TextString == groupName)     // нашли нужный блок с соотвествующим тегом
                                {
                                    def.Erase();
                                    break;
                                }
                            }
                        }
                    }

                    trExt.Commit();

                }
                // Сохраняем внешний файл
                extDb.SaveAs(dwgPath, DwgVersion.Current);
            }

            FillSignsGroups();

            if (RibbonInitializer.Instance.signsGroups.Items.Count > 0) RibbonInitializer.Instance.signsGroups.Current = RibbonInitializer.Instance.signsGroups.Items[0];
        }













        // возвращает текущий список групп знаков
        public static List<string> GetListOfSignGroups(Transaction tr, Database db, out BlockTableRecord btr)
        {
            var result = new List<string>();
            btr = null;

            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            // проверка наличия шаблона
            var btExt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            if (!btExt.Has("SIGN_GROUPS"))
            {
                ed.WriteMessage($"\n В БД не найден блок с группами знаков.");
                return result;
            }

            // получаем ID определения блока 
            var btrID = btExt["SIGN_GROUPS"];

            btr = (BlockTableRecord)tr.GetObject(btrID, OpenMode.ForWrite);

            foreach (ObjectId id in btr)    // проходимся по всем объектам и в список записываем только мтекст
            {
                var dbo = tr.GetObject(id, OpenMode.ForRead);
                if (dbo is MText mt) result.Add(mt.Text);
            }

            return result;
        }







        //  **************************************************************   TEMP  **************************************************************    //

        public static void CreateStandTemplateBlock()
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            const string blockName = "STAND_TEMPLATE"; // имя блока

            using (var lok = doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                //ObjectId btrId;
                if (bt.Has(blockName)) return;

                // Создаём пустое определение блока
                bt.UpgradeOpen();
                var btr = new BlockTableRecord { Name = blockName };
                bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                // Добавляем невидимые атрибуты 
                AddHiddenAttr(btr, tr, tag: "STAND", prompt: "Тег для идентификации ", defaultValue: "", 0);

                //AddHiddenAttr(btr, tr, tag: "STANDHANDLE", prompt: "Тег привязки знака к стойке", defaultValue: "", 2);

                tr.Commit();
            }
        }

        public static void CreateSignTemplateBlock()
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            const string blockName = "SIGN_TEMPLATE"; // имя блока

            using (var lok = doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                //ObjectId btrId;
                if (bt.Has(blockName)) return;

                // Создаём пустое определение блока
                bt.UpgradeOpen();
                var btr = new BlockTableRecord { Name = blockName };
                bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                // Добавляем невидимые атрибуты 
                AddHiddenAttr(btr, tr, tag: "SIGN", prompt: "Тег для идентификации ", defaultValue: "", 0);
                AddHiddenAttr(btr, tr, tag: "GROUP", prompt: "Группа знаков", defaultValue: "", 1);
                AddHiddenAttr(btr, tr, tag: "STANDHANDLE", prompt: "Тег привязки знака к стойке", defaultValue: "", 2);
                AddHiddenAttr(btr, tr, tag: "DOUBLED", prompt: "Два знака на одной стойке", defaultValue: "false", 3);

                tr.Commit();
            }
        }


        public static void DeleteTemplateBlocks()
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            HashSet<string> blockNames = new HashSet<string> { "SIGN_TEMPLATE", "STAND_TEMPLATE" };

            using (var lok = doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                foreach (var bl_name in blockNames)
                {
                    if (bt.Has(bl_name))
                    {
                        var btID = bt[bl_name];
                        var btr = (BlockTableRecord)tr.GetObject(btID, OpenMode.ForWrite);
                        if (!btr.IsErased) btr.Erase();
                    }
                }

                tr.Commit();
            }
        }



        // получить список блоков 
        public static List<T> GetListOfRefBlocks <T>(Dictionary<string, string> dictionary = null) 
            where T : IRefBlock, new()
        {
            var result = new List<T>();

            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            using (var lok = doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);     // таблица всех блоков 

                foreach (var blockId in bt)
                {
                    var btDef = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);
                    
                    // если это внешняя ссылка (attach/overlay) — пропускаем
                    if (btDef.IsFromExternalReference || btDef.IsFromOverlayReference) continue;

                    // ищем нужный tag
                    bool tagMatch = false;

                    foreach (var objId in btDef)
                    {
                        if (tr.GetObject(objId, OpenMode.ForRead) is AttributeDefinition at)
                        {
                            if (at.Tag.Equals(typeof(T).Name, StringComparison.OrdinalIgnoreCase))     // нашли нужный блок с соотвествующим тегом
                            {
                                tagMatch = true;
                                break;
                            }
                        }
                    }

                    // если определение блока подходит
                    if (tagMatch)
                    {
                        // получаем все вхождения блока
                        foreach (ObjectId btRefId in btDef.GetBlockReferenceIds(/*directOnly*/ true, /*includeErased*/ false))
                        {
                            var brRef = (BlockReference)tr.GetObject(btRefId, OpenMode.ForRead);
                           
                            // создаем экземпляр списка
                            var block = new T();

                            switch (block)
                            {
                                case Stand stand:
                                    
                                    //Обход атрибутов вставки
                                    foreach (ObjectId arId in brRef.AttributeCollection)
                                    {
                                        var ar = (AttributeReference)tr.GetObject(arId, OpenMode.ForRead);
                                        if (ar.Tag.Equals("ОСЬ", StringComparison.OrdinalIgnoreCase))
                                        {
                                            stand.AxisName = ar.TextString;  // наименование оси к которой привязана стойка
                                            break;
                                        }
                                    }

                                    stand.Handle = brRef.Handle.ToString();    // Handle вставки блока

                                    GetStandProperties(tr, db, stand, brRef.Position);  // Заполняем остальные данные

                                    result.Add(block);

                                break;

                                case Sign sign:

                                    //Обход атрибутов вставки
                                    foreach (ObjectId arId in brRef.AttributeCollection)
                                    {
                                        var ar = (AttributeReference)tr.GetObject(arId, OpenMode.ForRead);

                                        switch (ar.Tag)
                                        { 
                                            case var a when a.Equals("STANDHANDLE", StringComparison.OrdinalIgnoreCase):
                                                sign.StandHandle = ar.TextString;
                                                break;

                                            case var a when a.Equals("НОМЕР_ЗНАКА", StringComparison.OrdinalIgnoreCase):
                                                sign.Number = ar.TextString;
                                                break;

                                            case var a when a.Equals("DOUBLED", StringComparison.OrdinalIgnoreCase):
                                                sign.Doubled = ar.TextString == "True" ? "2" : "1";
                                                break;
                                        }
                                    }

                                    // параметры динамического блока
                                    foreach (DynamicBlockReferenceProperty prop in brRef.DynamicBlockReferencePropertyCollection)
                                    {
                                        if (prop.PropertyName.Equals("Типоразмер", StringComparison.OrdinalIgnoreCase))
                                        {
                                            sign.TypeSize = prop.Value.ToString();
                                        }
                                        if (prop.PropertyName.Equals("Наличие", StringComparison.OrdinalIgnoreCase))
                                        {
                                            sign.Existence = prop.Value.ToString();
                                        }
                                    }

                                    // вытягиваем имя из словаря
                                    if (dictionary.ContainsKey(sign.Number))
                                    {
                                        string nameOfSign = null;
                                        dictionary.TryGetValue(sign.Number, out nameOfSign);
                                        if (nameOfSign != null)
                                        {
                                            sign.Name = nameOfSign;
                                        }
                                    }
                                    else
                                    {
                                        sign.Name = "не определен";
                                    }
                                    result.Add(block);

                                break;
                            }
                        }
                    }
                }

                tr.Commit();
            }
            return result;
        }




        private static void GetStandProperties(Transaction tr, Database db, Stand stand, Point3d insertPoint)
        {
            Axis axis = TsoddHost.Current.axis.FirstOrDefault(a => a.Name == stand.AxisName);   // ось к которой привязанна стойка
            if (axis == null)
            {
                MessageBox.Show($"Ошибка. Не получилось найти ось {stand.AxisName} при создании объекта Stand (стойка)." +
                    $" Проверь блок в координатах {insertPoint} ");

                return;
            }

            Polyline axisPolyline = axis.AxisPoly;   // полилиния оси

            Point3d pkPoint = axisPolyline.GetClosestPointTo(insertPoint, false);   // точка ПК стойки на оси

            // считаем ПК
            double koef = db.Insunits == UnitsValue.Meters ? 1 : 0.001;     // коэффициент перевода в метры
            var distance = axisPolyline.GetDistAtPoint(pkPoint) * koef;   // расстояние от начала 
            if (axis.ReverseDirection) distance = axisPolyline.Length * koef - distance;    // если реверсивное направление оси
            distance = Math.Round(distance, 3) + axis.StartPK;

            stand.Distance = distance;  // расстояние от начала оси до точки ПК

            int pt_1 = (int)Math.Truncate(distance / 100);
            double pt_2 = Math.Round((distance - pt_1 * 100), 2);

            stand.PK = $"ПК {pt_1} + {pt_2}";  // пикет

            // положение слева/справа от оси (на оси)
            // если расстояние меньше чем указано в настройках, то считаем, что на оси
            double minDist = 1 / koef;
            double distToAxis = insertPoint.DistanceTo(pkPoint);
            if (distToAxis <= minDist)
            {
                stand.Side = "На оси";
            }
            else
            {
                try
                {
                    double vectDist = axis.ReverseDirection ? axisPolyline.GetDistAtPoint(pkPoint) + distToAxis :
                                                                axisPolyline.GetDistAtPoint(pkPoint) - distToAxis;

                    Point3d secondPointOnAxis = axisPolyline.GetPointAtDist(vectDist);

                    Vector3d vectorAxis = (pkPoint - secondPointOnAxis).GetNormal();
                    Vector3d vrctorPoint = (insertPoint - secondPointOnAxis).GetNormal();

                    var cross = vectorAxis.CrossProduct(vrctorPoint);

                    stand.Side = cross.Z < 0 ? "Справа от оси" : "Слева от оси";
                }
                catch
                {
                    stand.Side = "Не получилось определить";

                    MessageBox.Show($"Не удалось отсроить вектор для стойки при определении стороны относительно оси. " +
                        $"Проверь блок в координатах {insertPoint}");
                }

            }
        }












    }
}











