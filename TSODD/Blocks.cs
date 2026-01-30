
using TSODD;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsSystem;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using TSODD.forms;
using ACAD = Autodesk.AutoCAD.ApplicationServices;


namespace TSODD
{
    internal static class TsoddBlock
    {

        public static bool blockInsertFlag = true;

        // метод создания нового атрибута
        private static void AddAttributeToBlock(BlockTableRecord btr, Transaction tr, string tag, string prompt,
                                                string defaultValue, Point3d insertPoint, int numberOfAttr = 0,
                                                double height = 2.5, bool visible = false, bool inVisible = true, bool lockedPosition = true)
        {
            var ad = new AttributeDefinition
            {
                Position = insertPoint == Point3d.Origin ? new Point3d(0, -numberOfAttr * 2.5, 0) : insertPoint,
                Tag = tag.ToUpperInvariant(),
                Prompt = prompt ?? string.Empty,
                TextString = defaultValue ?? string.Empty,
                Height = height,
                Invisible = inVisible,
                Visible = visible,
                Constant = false,
                LockPositionInBlock = lockedPosition
            };

            btr.AppendEntity(ad);
            tr.AddNewlyCreatedDBObject(ad, true);
        }


        // метод изменения отрибутов блока
        public static void ChangeAttribute(Transaction tr, BlockReference br, string tag, string value)
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
                    if (!ar.Invisible)
                    {
                        // пользовательские настройки
                        var userOptions = JsonReader.LoadFromJson<Options>(FilesLocation.JsonOptionsPath);

                        // Получаем таблицу текстовых стилей
                        TextStyleTable textStyleTable = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);

                        // Перебираем все стили
                        foreach (ObjectId styleId in textStyleTable)
                        {
                            TextStyleTableRecord styleRecord = (TextStyleTableRecord)tr.GetObject(styleId, OpenMode.ForRead);
                            if (!string.IsNullOrEmpty(styleRecord.Name) && styleRecord.Name == userOptions.BlockNameTextStyle) ar.TextStyleId = styleId;
                        }

                        if (ad.Annotative == AnnotativeStates.True || btr.Annotative == AnnotativeStates.True || br.Annotative == AnnotativeStates.True)
                        {
                            ar.AddContext(cur);
                            ar.Height = userOptions.BlockNameTextHeight * cur.DrawingUnits;
                            ar.Position = new Point3d(ar.Position.X, ar.Position.Y - ar.Height / 2, 0);
                        }
                        else
                        {
                            ar.Height = userOptions.BlockNameTextHeight;
                            ar.Position = new Point3d(ar.Position.X, ar.Position.Y - ar.Height / 2, 0);
                        }
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


        // метод импортирования блоков с шаблона с настройками атрибутов
        private static void InsertBlockByNameWithTags(Point3d insertPoint, string name, Dictionary<string, string> tags)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            BlockTableRecord btr = null;
            ObjectIdCollection oldOrder = null;
            string hash_1 = "0";
            string hash_2 = "0";


            using (var blockDb = new Database(false, true))     // новая пустая база
            {
                blockDb.ReadDwgFile(FilesLocation.dwgBlocksPath, FileShare.Read, true, "");     // считываем в базу файл по указанному пути

                using (var trExt = blockDb.TransactionManager.StartTransaction())
                {
                    var btExt = (BlockTable)trExt.GetObject(blockDb.BlockTableId, OpenMode.ForRead);

                    if (!btExt.Has(name))
                    {
                        ed.WriteMessage($"\n Блок \"{name}\" не найден.");
                        blockInsertFlag = true;
                        return;
                    }

                    // блок в БД
                    var extBtrId = btExt[name];
                    BlockTableRecord extBtr = (BlockTableRecord)trExt.GetObject(extBtrId, OpenMode.ForRead);
                    DrawOrderTable drawOrderTableOld = (DrawOrderTable)trExt.GetObject(extBtr.DrawOrderTableId, OpenMode.ForRead);
                    oldOrder = drawOrderTableOld.GetFullDrawOrder(0);

                    // считаем его xэш
                    hash_1 = CalculateBlockHash(trExt, blockDb, extBtr);

                    // Импортируем определение блока в текущий чертёж
                    using (doc.LockDocument())
                    using (var tr = db.TransactionManager.StartTransaction())
                    {

                        // проверяем наличие блока
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                        if (bt.Has(name))
                        {
                            var blockId = bt[name];
                            btr = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForWrite);
                            hash_2 = CalculateBlockHash(tr, db, btr);
                        }

                        // если блоки разные, то перезаписываем блок
                        if (!hash_1.Equals(hash_2))
                        {
                            // Скопировать определение и все его зависимости (слои, стили и т.п.)
                            var ids = new ObjectIdCollection { extBtrId };
                            var mapping = new IdMapping();

                            // ownerId — это BlockTable текущей БД
                            blockDb.WblockCloneObjects(
                            ids,
                            db.BlockTableId,
                            mapping,
                            DuplicateRecordCloning.Replace, // Ignore=оставить существующий
                            false
                            );

                            // Получаем ObjectId склонированного (или существующего) определения в текущей БД
                            var pair = mapping[extBtrId];
                            var destBtrId = pair.Value;

                            btr = (BlockTableRecord)tr.GetObject(destBtrId, OpenMode.ForWrite);

                            // Сихронизируем порядок отрисовки элементов в блоках
                            SyncDrawOrder(oldOrder, btr, mapping, tr);
                        }

                        // аннатотивность блока
                        var userOptions = JsonReader.LoadFromJson<Options>(FilesLocation.JsonOptionsPath);
                        btr.Annotative = userOptions.BlocksAnnotativeState ? AnnotativeStates.True : AnnotativeStates.False;


                        blockInsertFlag = false;

                        //var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        var br = new BlockReference(insertPoint, btr.Id);

                        ms.AppendEntity(br);
                        tr.AddNewlyCreatedDBObject(br, true);

                        // масштаб аннотаций
                        if (userOptions.BlocksAnnotativeState)
                        {
                            var ocm = db.ObjectContextManager;
                            var occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");
                            // Текущий масштаб CANNOSCALE
                            var curSC = occ.CurrentContext as AnnotationScale;
                            if (curSC != null && !br.HasContext(curSC))

                            br.AddContext(curSC);   // добавляем масштаб в блок
                        }

                        // заполняем атрибуты
                        SetAttrValues(br, tr, tags);

                        br.RecordGraphicsModified(true);

                        blockInsertFlag = true;  // можно опять вставлять блоки

                        tr.Commit();
                    }
                    trExt.Commit();
                }
            }


            // метод для подсчета Hash блока
            string CalculateBlockHash(Transaction transaction, Database database,  BlockTableRecord block)
            {
                using (var stream = new MemoryStream())
                {
                    using (var writer = new BinaryWriter(stream))
                    {
                        int entityCount = 0;
                        double max_X = double.MinValue;
                        double max_Y = double.MinValue;
                        double min_X = double.MaxValue;
                        double min_Y = double.MaxValue;

                        foreach (var objId in block)
                        {
                            Entity ent = (Entity)transaction.GetObject(objId, OpenMode.ForRead);
                            if (ent.Bounds != null)
                            {
                                if (Math.Round(ent.Bounds.Value.MaxPoint.X, 3) > max_X) max_X = ent.Bounds.Value.MaxPoint.X;
                                if (Math.Round(ent.Bounds.Value.MaxPoint.Y, 3) > max_Y) max_Y = ent.Bounds.Value.MaxPoint.Y;
                                if (Math.Round(ent.Bounds.Value.MinPoint.X, 3) < min_X) min_X = ent.Bounds.Value.MinPoint.X;
                                if (Math.Round(ent.Bounds.Value.MinPoint.Y, 3) < min_Y) min_Y = ent.Bounds.Value.MinPoint.Y;
                            }
                            entityCount += 1;
                        }

                        // границы блока
                        writer.Write(max_X);
                        writer.Write(max_Y);
                        writer.Write(min_X);
                        writer.Write(min_Y);
                        
                        // кол-во элементов
                        writer.Write(entityCount);

                        using (var md5 = System.Security.Cryptography.MD5.Create())
                        {
                            byte[] hash = md5.ComputeHash(stream.ToArray());
                            return BitConverter.ToString(hash).Replace("-", "").ToLower();
                        }

                    }
                }
            }





        }

        // метод вставки блока стойки
        public static void InsertStandOrMarkBlock(string name, bool stand)
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


            // запоминаем последний выбранный блок
            if (stand) TsoddHost.Current.currentStandBlock = name;
            if (!stand) TsoddHost.Current.currentMarkBlock = name;
        }


        // метод вставки блока знака
        public static void InsertSignBlock(string name)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // соберем все блоки стоек в отдельный словарь где будут точки вставки и ID блока
            Dictionary<ObjectId, Point3d> standBlocks = new Dictionary<ObjectId, Point3d>();
            var userOptions = JsonReader.LoadFromJson<Options>(FilesLocation.JsonOptionsPath);


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

                double maxRadius = db.Insunits == UnitsValue.Meters ? userOptions.BlockBindRadius : userOptions.BlockBindRadius * 1000;    // максимальный радиус поиска

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
                                var bref = (BlockReference)tr.GetObject(lastHighlightedId, OpenMode.ForRead);
                                var tagList = new Dictionary<string, string> { ["STANDHANDLE"] = bref.Handle.ToString() };
                                // вставляем блок
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
                                using (doc.LockDocument())
                                {
                                    using (var tr = db.TransactionManager.StartTransaction())
                                    {
                                        var bref = (BlockReference)tr.GetObject(lastHighlightedId, OpenMode.ForRead);
                                        var tagList = new Dictionary<string, string> { ["STANDHANDLE"] = bref.Handle.ToString() };

                                        // вставляем блок
                                        InsertBlockByNameWithTags(ppr.Value, name, tagList);
                                        tr.Commit();

                                    }
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

                    // запоминаем последний выбранный блок
                    TsoddHost.Current.currentSignBlock = name;
                }
            }
        }


        //  **************************************************************   ADD / DELETE  BLOCKS   **************************************************************    //


        // метод загрузки блока в базу
        public static void AddBlockToBD(string templateName, BlockTableRecord intBtr, string intBlockName, string intBlockNumber,
            ObservableCollection<DataGridValue> collection, List<ObjectId> deleteObjects,
            Extents3d extents, Bitmap blockIcon, bool singleSign, string groupName)
        {
            RibbonInitializer.Instance.readyToDeleteEntity = false; // запрещаем анализировать объекты при удалении

            ObjectIdCollection idsToCloneOriginalBlock = new ObjectIdCollection();
            ObjectIdCollection idsToCloneTemplateBlock = new ObjectIdCollection();

            ObjectIdCollection oldOrder = null;

            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            Database templateDb = null;
            Database blocksDb = null;
            ObjectId templateId;
            ObjectId blockBtrId;
            BlockTableRecord blockBtr;

            try
            {
                using (doc.LockDocument())
                {
                    // берем данные с исходного блока
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        // просто возьмем все элементы исходного блока
                        idsToCloneOriginalBlock = new ObjectIdCollection { intBtr.ObjectId };

                        // Сохраняем draw order, пока Table ещё в транзакции
                        if (!intBtr.DrawOrderTableId.IsNull)
                        {
                            var drawOrderTableOld = (DrawOrderTable)tr.GetObject(intBtr.DrawOrderTableId, OpenMode.ForRead);
                            oldOrder = drawOrderTableOld.GetFullDrawOrder(0);
                        }

                    }
                }

                // копируем шаблон из файла с блоками, шаманим с ним и сохраняем под новым именем
                // Открываем DWG с шаблонами
                templateDb = new Database(false, true);
                templateDb.ReadDwgFile(FilesLocation.dwgTemplatePath, FileShare.None, false, "");
                templateDb.CloseInput(true);

                // Открываем внешний DWG с блоками
                blocksDb = new Database(false, true);
                blocksDb.ReadDwgFile(FilesLocation.dwgBlocksPath, FileShare.None, false, "");
                blocksDb.CloseInput(true);

                // Находим шаблон в внешней БД
                using (var trForTemplate = templateDb.TransactionManager.StartTransaction())
                {
                    // проверка наличия шаблона
                    var btExt = (BlockTable)trForTemplate.GetObject(templateDb.BlockTableId, OpenMode.ForRead);
                    if (!btExt.Has(templateName))
                    {
                        ed.WriteMessage($"\nВ БД не найден шаблон \"{templateName}\".");
                        RibbonInitializer.Instance.readyToDeleteEntity = true; // разрешаем анализировать объекты при удалении
                        return;
                    }

                    templateId = btExt[templateName];   // id блока шаблона
                    var templateBtr = (BlockTableRecord)trForTemplate.GetObject(templateId, OpenMode.ForRead);

                    //  копируем все элементы шаблона в коллекцию
                    foreach (ObjectId id in templateBtr) idsToCloneTemplateBlock.Add(id);

                    trForTemplate.Commit();
                }

                // проверяем наличие блока с таким же именем
                using (var trForBlock = blocksDb.TransactionManager.StartTransaction())
                {
                    var btExt = (BlockTable)trForBlock.GetObject(blocksDb.BlockTableId, OpenMode.ForRead);

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
                            trForBlock.Abort();
                            RibbonInitializer.Instance.readyToDeleteEntity = true; // разрешаем анализировать объекты при удалении
                            return;
                        }
                    }

                    // Копируем исходный блок в чертеж БД с блоками
                    var blockMap = new IdMapping();
                    blocksDb.WblockCloneObjects(
                    idsToCloneOriginalBlock,
                    blocksDb.BlockTableId,
                    blockMap,
                    DuplicateRecordCloning.Replace, // создаст копию записи блока
                    false
                    );

                    // Получаем id клона 
                    blockBtrId = ((IdPair)blockMap[intBtr.ObjectId]).Value;

                    // Переименовываем и настраиваем блок
                    blockBtr = (BlockTableRecord)trForBlock.GetObject(blockBtrId, OpenMode.ForWrite);

                    // Сихронизируем порядок отрисовки элементов в блоках
                    SyncDrawOrder(oldOrder, blockBtr, blockMap, trForBlock);


                    HashSet<string> dublicateValues = new HashSet<string> { "STAND", "ОСЬ", "SIGN", "GROUP", "STANDHANDLE", "DOUBLED",
                                                                             "НОМЕР_ЗНАКА", "TYPESIZE", "SIGNEXISTENCE", "MARK", "MATERIAL","MARKEXISTENCE" };
                    // чистим блок от ненужных атрибутов и других элементов
                    foreach (ObjectId id in intBtr)
                    {
                        ObjectId cloneId = blockMap[id].Value;
                        if (cloneId == null) continue;

                        Entity ent = (Entity)trForBlock.GetObject(cloneId, OpenMode.ForRead);

                        // если id был передан в коллекции deleteObjects
                        if (deleteObjects.Any(i => i == id))
                        {
                            ent.UpgradeOpen();
                            ent.Erase();
                            continue;
                        }

                        // если пользователь не захотел добавлять этот атрибут
                        var item = collection.FirstOrDefault(i => i.ID == id);
                        if (item != null && item.Value == false)
                        {
                            ent.UpgradeOpen();
                            ent.Erase();
                            continue;
                        }

                        if (ent is AttributeDefinition attr)
                        {
                            // если есть атрибуты из шаблонов
                            if (dublicateValues.Contains(attr.Tag))
                            {
                                attr.UpgradeOpen();
                                attr.Erase();
                                continue;
                            }
                        }
                    }


                    if (!btExt.Has(intBlockName)) blockBtr.Name = intBlockName;
                    blockBtr.PreviewIcon = blockIcon;  // меняем previewIcon у блока
                    blockBtr.Annotative = AnnotativeStates.True;
                    blockBtr.Name = intBlockName;

                    trForBlock.Commit();

                }

                // Добавляем атрибуты из блока шаблона в скопированный блок
                var templateMap = new IdMapping();
                blocksDb.WblockCloneObjects(
                idsToCloneTemplateBlock,
                blockBtrId,
                templateMap,
                DuplicateRecordCloning.Ignore, // создаст копию записи блока
                false
                );


                // настраиваем блок в зависимости от его типа
                switch (templateName)
                {
                    case "STAND_TAMPLATE":

                        break;

                    case "SIGN_TEMPLATE":
                        using (var tr = blocksDb.TransactionManager.StartTransaction())
                        {
                            // меняем атрибуты для знака
                            foreach (ObjectId id in blockBtr)
                            {

                                if (tr.GetObject(id, OpenMode.ForWrite) is AttributeDefinition ad && !ad.Constant)
                                {

                                    if (ad.Tag.Equals("НОМЕР_ЗНАКА", StringComparison.OrdinalIgnoreCase))
                                    {
                                        var entityWidth = (extents.MaxPoint.X - extents.MinPoint.X);
                                        var entityHeight = (extents.MaxPoint.Y - extents.MinPoint.Y);

                                        double offset = Math.Max(entityWidth, entityHeight) / 20;
                                        ad.TextString = intBlockNumber;
                                        ad.Position = new Point3d(extents.MinPoint.X + entityWidth + offset,
                                                                    extents.MinPoint.Y + entityHeight / 2, 0);
                                    }

                                    if (ad.Tag.Equals("GROUP", StringComparison.OrdinalIgnoreCase))
                                    {
                                        if (TsoddHost.Current.currentSignGroup == null) MessageBox.Show("Ошибка выбора текущей группы знаков");
                                        ad.TextString = groupName; // привязываем к группе знаков
                                    }

                                    if (ad.Tag.Equals("DOUBLED", StringComparison.OrdinalIgnoreCase))
                                    {
                                        ad.TextString = (!singleSign).ToString();
                                    }
                                }
                            }
                            tr.Commit();
                        }
                        break;

                    case "MARK_TEMPLATE":
                        // меняем атрибуты для знака
                        foreach (ObjectId id in blockBtr)
                        {
                            using (var tr = blocksDb.TransactionManager.StartTransaction())
                            {
                                if (tr.GetObject(id, OpenMode.ForWrite) is AttributeDefinition ad && !ad.Constant)
                                {
                                    if (ad.Tag.Equals("GROUP", StringComparison.OrdinalIgnoreCase))
                                    {
                                        if (TsoddHost.Current.currentMarkGroup == null) MessageBox.Show("Ошибка выбора текущей группы разметки");
                                        ad.TextString = groupName; // привязываем к группе 
                                    }
                                    if (ad.Tag.Equals("NUMBER", StringComparison.OrdinalIgnoreCase))
                                    {
                                        ad.TextString = intBlockNumber; // номер разметки
                                    }

                                }
                                tr.Commit();
                            }
                        }

                        break;

                }

                // Сохраняем внешний файл
                blocksDb.SaveAs(FilesLocation.dwgBlocksPath, DwgVersion.Current);

            }
            catch
            {
                ed.WriteMessage($"\n Блок \"{intBlockName}\" не получилось добавить в базу. \n");
                return;
            }
            finally
            {
                RibbonInitializer.Instance.readyToDeleteEntity = true; // разрешаем анализировать объекты при удалении
                templateDb?.Dispose();
                blocksDb?.Dispose();
            }

            ed.WriteMessage($"\n Блок \"{intBlockName}\" успешно добавлен в базу. \n");
        }



        public static void SyncDrawOrder(ObjectIdCollection oldOrder, BlockTableRecord newBtr, IdMapping map, Transaction tr)
        {
            // применяем порядок отрисовки примитивов
            if (oldOrder != null && oldOrder.Count > 0)
            {
                DrawOrderTable drawOrderTableNew = tr.GetObject(newBtr.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;   // новая таблица DrawOrderTable
                ObjectIdCollection newOrder = new ObjectIdCollection();                 // новый порядок

                // перебор всех элементов в старом порядке
                foreach (ObjectId id in oldOrder)
                {
                    if (!map.Contains(id)) continue; // не нашли такой id
                    IdPair pair = map[id];

                    if (!pair.IsCloned || pair.Value.IsNull)    //объект не был скопирован в новый блок или новый Id отсутствует
                        continue;

                    newOrder.Add(pair.Value);
                }

                drawOrderTableNew.SetRelativeDrawOrder(newOrder);
            }
        }



        public static void DeleteBlockFromBD(string blockName)
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (var extDb = new Database(false, true))
            {
                // Открываем внешний DWG для записи
                extDb.ReadDwgFile(FilesLocation.dwgBlocksPath, FileShare.None, false, "");
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
       
                        // получаем ID определения блока 
                        var btrID = btExt[blockName];

                        var btr = (BlockTableRecord)trExt.GetObject(btrID, OpenMode.ForWrite);
                        if (!btr.IsErased) btr.Erase();
                    }

                    trExt.Commit();
                }

                // Сохраняем внешний файл
                extDb.SaveAs(FilesLocation.dwgBlocksPath, DwgVersion.Current);
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



        public static Bitmap GetBlockBitmap(BlockTableRecord btr, byte px, out Extents3d extents)
        {
            Bitmap bmpResult = null;
            Bitmap bmpBackground = null;
            Bitmap bmpText = null;

            bmpBackground = GetBlockBitmapPart(btr, px, false, out extents); // делаем снимок блока - все кроме текста
            bmpText = GetBlockBitmapPart(btr, px, true, out extents); // делаем снимок блока - только текст

            bmpResult = JoinBitmapForBlockIcon(bmpBackground, bmpText);

            bmpBackground.Dispose();
            bmpText.Dispose();

            return bmpResult;
        }


        private static Bitmap GetBlockBitmapPart(BlockTableRecord btr, byte px, bool onlyText, out Extents3d extents)
        {

            var db = btr.Id.Database ?? Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database; // получаем БД объекта
            //var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.GetDocument(db);

            Bitmap bmp = null;
            bool hasExt = false;
            extents = new Extents3d();

            Manager manager = ACAD.Application.DocumentManager.MdiActiveDocument.GraphicsManager;

            var descriptor = new KernelDescriptor();
            descriptor.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"));

            var kernel = Manager.AcquireGraphicsKernel(descriptor);
            var device = manager.CreateAutoCADOffScreenDevice(kernel);
            var model = manager.CreateAutoCADModel(kernel);

            try
            {
                using (device)
                {
                    device.OnSize(new System.Drawing.Size(px, px));

                    using (model)
                    {
                        var view = new Autodesk.AutoCAD.GraphicsSystem.View();
                        device.Add(view);

                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            // 1. Берём порядок из DrawOrderTable
                            ObjectIdCollection order = null;
                            if (!btr.DrawOrderTableId.IsNull)
                            {
                                var dot = (DrawOrderTable)tr.GetObject(btr.DrawOrderTableId, OpenMode.ForRead);
                                order = dot.GetFullDrawOrder(0);
                            }

                            // Если таблицы нет – просто по порядку в BTR
                            IEnumerable<ObjectId> ids = order != null && order.Count > 0 ? order.Cast<ObjectId>() : btr.Cast<ObjectId>();

                            // Пребираем все элементы по drawOrder и добаваляем на вид
                            foreach (ObjectId id in ids)
                            {
                                Entity ent = (Entity)tr.GetObject(id, OpenMode.ForRead).Clone();

                                try
                                {
                                    Extents3d ext = ent.GeometricExtents;
                                    extents.AddExtents(ext);
                                    hasExt = true;
                                }
                                catch { }

                                // если элемент текст или атрибут, то настраиваем их
                                bool isTextLike = ent is DBText || ent is MText || ent is AttributeDefinition || ent is AttributeReference;
                                if (isTextLike) ent.ColorIndex = 3;
                                if (ent is AttributeDefinition attr)
                                {
                                    attr.Tag = attr.TextString;
                                    attr.AdjustAlignment(db);
                                }

                                // если выводим только текст и элемент текст или выводим все кроме текста и элемент не текст
                                if (onlyText == isTextLike) view.Add(ent, model);
                            }

                            tr.Commit();
                        }

                        // находим границы
                        if (hasExt)
                        {
                            var minX = extents.MinPoint.X;
                            var minY = extents.MinPoint.Y;
                            var entityWidth = Math.Round(extents.MaxPoint.X - minX, 2);
                            var entityHeight = Math.Round(extents.MaxPoint.Y - minY, 2);
                            var center = new Point3d((extents.MaxPoint.X + extents.MinPoint.X) * 0.5, (extents.MaxPoint.Y + extents.MinPoint.Y) * 0.5, 0.0);

                            try
                            {   // настраиваем зум вида
                                view.SetView(center, center, view.UpVector, entityWidth, entityHeight);
                                view.ZoomExtents(extents.MinPoint, extents.MaxPoint);
                            }
                            catch { }
                        }

                        System.Drawing.Rectangle rect = new System.Drawing.Rectangle(0, 0, px, px);
                        bmp = view.GetSnapshot(rect);
                    }

                }
            }
            finally
            {
                Manager.ReleaseGraphicsKernel(kernel);
            }

            return bmp;
        }



        private static Bitmap JoinBitmapForBlockIcon(Bitmap bmpBackground, Bitmap bmpText)
        {
            Bitmap bitmap = null;
            if (bmpBackground == null || bmpText == null) return bitmap;
            if (bmpBackground.Width != bmpText.Width || bmpBackground.Height != bmpText.Height) return bitmap;  // почему-то не совпадают размеры bitmap

            bitmap = new Bitmap(bmpBackground);

            var backColor = bmpText.GetPixel(0, 0); // цвет фона у картинки с текстом

            for (int y = 0; y < bmpText.Height; y++)
            {
                for (int x = 0; x < bmpText.Width; x++)
                {
                    var p = bmpText.GetPixel(x, y);

                    // если пиксель не фон — кладём его поверх
                    if (!ColorsEqual(p, backColor))
                    {
                        bitmap.SetPixel(x, y, p);
                    }
                }
            }

            // внутренний метод сравнивает значение цвета пикселя с фоном с учетом погршности tol
            bool ColorsEqual(System.Drawing.Color a, System.Drawing.Color b, int tol = 5)
            {
                return Math.Abs(a.R - b.R) <= tol &&
                       Math.Abs(a.G - b.G) <= tol &&
                       Math.Abs(a.B - b.B) <= tol;
            }

            return bitmap;
        }



        // метод получает список имен блоков и их previewIcon, которые соответсвуют тегу (для заполнения RibbonSplitButton)
        public static List<(string name, BitmapSource img)> GetListOfBlocks(string typeTag, string group)
        {
            List<(string name, BitmapSource img)> result = new List<(string, BitmapSource)>(); // возвращаемый список блоков (имя и картинка привью)

            using (var blockDb = new Database(false, true))     // новая пустая база
            {
                blockDb.ReadDwgFile(FilesLocation.dwgBlocksPath, FileShare.Read, true, "");     // считываем в базу файл по указанному пути

                using (var tr = blockDb.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(blockDb.BlockTableId, OpenMode.ForRead);

                    foreach (var blockID in bt)
                    {
                        var def = (BlockTableRecord)tr.GetObject(blockID, OpenMode.ForRead);

                        //if (def.Name.Contains("_TEMPLATE")) continue;

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

            Axis selectedAxis = RibbonInitializer.Instance.SelectAxis();  // готовая команда вернет выбранную ось
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

        //  **************************************************************   РАБОТА С РАЗМЕТКОЙ   **************************************************************    //

        public static void CreateUserMarkBlock()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var objectIdList = RibbonInitializer.Instance.GetAutoCadSelectionObjectsId(new List<string>());
            if (objectIdList == null || objectIdList.Count == 0)
            {
                ed.WriteMessage(" \n Не выбраны объекты. Сначала выберите объекты, которые следует добавить в пользовательский блок разметки. \n");
                return; // ничего не выбрали
            }

            // запрашиваем точки
            List<Point3d> pointsForPK = new List<Point3d>();
            PromptPointOptions ppo = new PromptPointOptions($"\n Выберите точку, для подсчета ПК (Esc - продолжить/выход): ");
            ppo.AllowNone = true;
            int numerator = 0;

            while (true)
            {
                numerator++;
                ppo.Message = $"\n Выберите точку {numerator}, для подсчета ПК (Esc/Enter - выход): ";
                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status != PromptStatus.OK) break;
                pointsForPK.Add(ppr.Value);
            }

            if (pointsForPK.Count == 0) //  в итоге не выбрали точки для ПК
            {
                ed.WriteMessage("\n Ошибка создания пользовательского блока разметки");
                return;
            }

            PromptStringOptions pso = new PromptStringOptions("Введите номер разметки");
            pso.AllowSpaces = false;
            PromptResult psr = ed.GetString(pso);
            if (psr.Status != PromptStatus.OK) //  в итоге не выбрали точки для ПК
            {
                ed.WriteMessage("\n Ошибка создания пользовательского блока разметки");
                return;
            }


            // создаем определение блока 
            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                string newNameOfBlock = GenerateUniqueBlockName($"{psr.StringResult}_UserMarkBlock", blockTable);

                var btr = new BlockTableRecord();
                btr.Name = newNameOfBlock;
                btr.Origin = pointsForPK[0]; // Устанавливаем точку вставки блока
                btr.Annotative = AnnotativeStates.False;
                btr.Units = db.Insunits;

                blockTable.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                // добавляе Entities из выбора objectIdList
                List<Entity> entities = new List<Entity>();
                RibbonInitializer.Instance.readyToDeleteEntity = false;
                foreach (var id in objectIdList)
                {
                    Entity ent = (Entity)tr.GetObject(id, OpenMode.ForWrite);
                    Entity clonedEnt = (Entity)ent.Clone();

                    // добавляем в блок 
                    btr.AppendEntity(clonedEnt);
                    tr.AddNewlyCreatedDBObject(clonedEnt, true);

                    // удаляем исходные entity
                    ent.Erase();
                }

                // добавляем атрибуты для ПК
                numerator = 0;
                foreach (var point in pointsForPK)
                {
                    numerator++;
                    AddAttributeToBlock(btr, tr, $"PK_VAL_{numerator}", "", "", point, 0, 5, true, true);
                }
                // добавляем атрибут привязки к оси
                AddAttributeToBlock(btr, tr, tag: "ОСЬ", prompt: "Тег привязки к оси",
                                    defaultValue: $"{TsoddHost.Current.currentAxis.Name}",
                                    pointsForPK[0], 0, 2.5, true, true);
                AddAttributeToBlock(btr, tr, tag: "MARK", prompt: "Тег для идентификации ", defaultValue: "", Point3d.Origin,1);
                AddAttributeToBlock(btr, tr, tag: "NUMBER", prompt: "Тег для идентификации ", defaultValue: $"{psr.StringResult}", Point3d.Origin,2);
                AddAttributeToBlock(btr, tr, tag: "MATERIAL", prompt: "Материал", defaultValue: "Термопластик", Point3d.Origin, 3);
                AddAttributeToBlock(btr, tr, tag: "MARKEXISTENCE", prompt: "Наличие", defaultValue: "Требуется нанести", Point3d.Origin, 4);


                // делаем вставку блока
                BlockReference br = new BlockReference(pointsForPK[0], btr.ObjectId);
                var ms = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                ms.AppendEntity(br);
                tr.AddNewlyCreatedDBObject(br, true);

                SetAttrValues(br, tr, new Dictionary<string, string> { ["ОСЬ"] = TsoddHost.Current.currentAxis.Name });

                tr.Commit();
                RibbonInitializer.Instance.readyToDeleteEntity = true;

                ed.WriteMessage($"Пользовательский блок создан. Имя блока: {newNameOfBlock}");
            }

            string GenerateUniqueBlockName(string baseName, BlockTable bt)
            {
                string name = baseName;
                int counter = 1;

                while (bt.Has(name))
                {
                    name = $"{baseName}_{counter}";
                    counter++;
                }
                return name;
            }
        }



        //  **************************************************************   РАБОТА С ГРУППАМИ   **************************************************************    //

        // предвыбор групп при открытии файла
        public static void PreSelectOfGroups()
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (var extDb = new Database(false, true))
            {
                // Открываем внешний DWG для записи
                extDb.ReadDwgFile(FilesLocation.dwgBlocksPath, FileShare.None, false, "");
                extDb.CloseInput(true);

                using (var trExt = extDb.TransactionManager.StartTransaction())
                {
                    // список групп в блоке 
                    var groupListSign = GetListGroups("SIGN", trExt, extDb, out _);
                    var groupListMark = GetListGroups("MARK", trExt, extDb, out _);

                    if (groupListSign.Count > 0) TsoddHost.Current.currentSignGroup = groupListSign[0];
                    if (groupListMark.Count > 0) TsoddHost.Current.currentSignGroup = groupListMark[0];

                    trExt.Commit();

                }
            }
        }


        // добавляет группу в шаблон
        public static void AddGroupToBD(string template, string groupName)
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (var extDb = new Database(false, true))
            {
                // Открываем внешний DWG для записи
                extDb.ReadDwgFile(FilesLocation.dwgBlocksPath, FileShare.None, false, "");
                extDb.CloseInput(true);

                using (var trExt = extDb.TransactionManager.StartTransaction())
                {

                    // список групп в блоке знаков
                    var groupList = GetListGroups(template, trExt, extDb, out BlockTableRecord btr);

                    if (btr == null)    // если по какой-то причине у нас нет определения блока
                    {
                        MessageBox.Show($" \n Ошибка. Не найден блок \"{template}\" ");
                        return;
                    }

                    // проверяем нет ли дубликатов
                    var dublicate = groupList.Any(g => string.Equals(g, groupName, StringComparison.OrdinalIgnoreCase));

                    if (dublicate)
                    {
                        MessageBox.Show($" \n Группа с именем \"{groupName}\" уже существует");
                        return;
                    }

                    MText mText = new MText();
                    mText.Contents = groupName;

                    var id = btr.AppendEntity(mText);
                    trExt.AddNewlyCreatedDBObject(mText, true);

                    trExt.Commit();
                }

                // Сохраняем внешний файл,
                extDb.SaveAs(FilesLocation.dwgBlocksPath, DwgVersion.Current);
            }
        }


        // удаляет группу знаков из шаблона
        public static void DeleteGroupFromBD(string template, string groupName)
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            PromptKeywordOptions pso = new PromptKeywordOptions($"\n Вы точно хотите удалить группу \"{groupName}\" и входящие в нее блоки из БД? [Да/Нет]: ");
            pso.Keywords.Add("Да");
            pso.Keywords.Add("Нет");
            pso.AllowArbitraryInput = false;

            PromptResult pres = ed.GetKeywords(pso);

            if (pres.StringResult != "Да")  // выходим, если не хотим удалять группу
            {
                MessageBox.Show($" \n Ошибка удаления группы ...");
                return;
            }

            List<string> groupList = new List<string>();

            // Если таки удаляем группу и знаки
            using (var extDb = new Database(false, true))
            {
                // Открываем внешний DWG для записи
                extDb.ReadDwgFile(FilesLocation.dwgBlocksPath, FileShare.None, false, "");
                extDb.CloseInput(true);

                using (var trExt = extDb.TransactionManager.StartTransaction())
                {
                    // список групп в блоке знаков
                    groupList = GetListGroups(template, trExt, extDb, out BlockTableRecord btr);

                    if (btr == null)    // если по какой-то причине у нас нет определения блока
                    {
                        MessageBox.Show($" \n Ошибка. Не найден блок \"{template}\" ");
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
                                groupList.Remove(groupName);
                            }
                        }
                    }

                    // удаляем все блоки, которые входили в эту группу
                    var bt = (BlockTable)trExt.GetObject(extDb.BlockTableId, OpenMode.ForRead);

                    foreach (var blockID in bt)
                    {
                        var def = (BlockTableRecord)trExt.GetObject(blockID, OpenMode.ForWrite);

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
                extDb.SaveAs(FilesLocation.dwgBlocksPath, DwgVersion.Current);
            }

            if (groupList.Count > 0)
            {
                if (template == "SIGN") TsoddHost.Current.currentSignGroup = groupList[0];
                if (template == "MARK") TsoddHost.Current.currentMarkGroup = groupList[0];
            }

        }

        // возвращает текущий список групп 
        public static List<string> GetListGroups(string templateName, Transaction tr, Database db)
        {
            return GetListGroups(templateName, tr, db, out _);
        }

        public static List<string> GetListGroups(string templateName, Transaction tr, Database db, out BlockTableRecord btr)
        {
            var result = new List<string>();
            btr = null;

            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            // проверка наличия шаблона
            var btExt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            if (!btExt.Has(templateName))
            {
                ed.WriteMessage($"\n В БД не найден блок с группами \'{templateName}\'.");
                return result;
            }

            // получаем ID определения блока 
            var btrID = btExt[templateName];

            btr = (BlockTableRecord)tr.GetObject(btrID, OpenMode.ForWrite);

            foreach (ObjectId id in btr)    // проходимся по всем объектам и в список записываем только мтекст
            {
                var dbo = tr.GetObject(id, OpenMode.ForRead);
                if (dbo is MText mt) result.Add(mt.Text);
            }

            return result;
        }


        //  **************************************************************   TEMP  **************************************************************    //

        private static void CreateStandTemplateBlock()
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            const string blockName = "STAND_TEMPLATE"; // имя блока

            using (var lok = doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                //ObjectId brId;
                if (bt.Has(blockName)) return;

                // Создаём пустое определение блока
                bt.UpgradeOpen();
                var btr = new BlockTableRecord { Name = blockName };
                bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                // Добавляем невидимые атрибуты 
                AddAttributeToBlock(btr, tr, tag: "STAND", prompt: "Тег для идентификации ", defaultValue: "", Point3d.Origin, 0);
                AddAttributeToBlock(btr, tr, tag: "ОСЬ", prompt: "Тег привязки стойки к оси", defaultValue: "", Point3d.Origin, 1, 2.5, true);

                tr.Commit();
            }
        }


        private static void CreateSignTemplateBlock()
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            const string blockName = "SIGN_TEMPLATE"; // имя блока

            using (var lok = doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                //ObjectId brId;
                if (bt.Has(blockName)) return;

                // Создаём пустое определение блока
                bt.UpgradeOpen();
                var btr = new BlockTableRecord { Name = blockName };
                bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                // Добавляем невидимые атрибуты 
                AddAttributeToBlock(btr, tr, tag: "SIGN", prompt: "Тег для идентификации ", defaultValue: "", Point3d.Origin, 0);
                AddAttributeToBlock(btr, tr, tag: "GROUP", prompt: "Группа", defaultValue: "", Point3d.Origin, 1);
                AddAttributeToBlock(btr, tr, tag: "STANDHANDLE", prompt: "Тег привязки знака к стойке", defaultValue: "", Point3d.Origin, 2);
                AddAttributeToBlock(btr, tr, tag: "DOUBLED", prompt: "Два знака на одной стойке", defaultValue: "false", Point3d.Origin, 3);
                AddAttributeToBlock(btr, tr, tag: "НОМЕР_ЗНАКА", prompt: "Номер знака", defaultValue: "", Point3d.Origin, 4, 2.5, true, false, false);
                AddAttributeToBlock(btr, tr, tag: "TYPESIZE", prompt: "Типоразмер", defaultValue: "I", Point3d.Origin, 3);
                AddAttributeToBlock(btr, tr, tag: "SIGNEXISTENCE", prompt: "Наличие", defaultValue: "Необходимо установить", Point3d.Origin, 4);
                tr.Commit();
            }
        }

        private static void CreateMarkTemplateBlock()
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            const string blockName = "MARK_TEMPLATE"; // имя блока

            using (var lok = doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                //ObjectId brId;
                if (bt.Has(blockName)) return;

                // Создаём пустое определение блока
                bt.UpgradeOpen();
                var btr = new BlockTableRecord { Name = blockName };
                bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);

                // Добавляем невидимые атрибуты 
                AddAttributeToBlock(btr, tr, tag: "MARK", prompt: "Тег для идентификации ", defaultValue: "", Point3d.Origin, 0);
                AddAttributeToBlock(btr, tr, tag: "ОСЬ", prompt: "Тег привязки стойки к оси", defaultValue: "", Point3d.Origin, 1, 2.5, true);
                AddAttributeToBlock(btr, tr, tag: "NUMBER", prompt: "Номер", defaultValue: "", Point3d.Origin, 2);
                AddAttributeToBlock(btr, tr, tag: "GROUP", prompt: "Группа", defaultValue: "", Point3d.Origin, 3);
                AddAttributeToBlock(btr, tr, tag: "MATERIAL", prompt: "Материал", defaultValue: "Термопластик", Point3d.Origin, 4);
                AddAttributeToBlock(btr, tr, tag: "MARKEXISTENCE", prompt: "Наличие", defaultValue: "Требуется нанести", Point3d.Origin, 5);

                tr.Commit();
            }
        }


        private static void CreateSignGroupsTemplateBlock()
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            const string blockName = "SIGN_GROUPS"; // имя блока

            using (var lok = doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                //ObjectId brId;
                if (bt.Has(blockName)) return;

                // Создаём пустое определение блока
                bt.UpgradeOpen();
                var btr = new BlockTableRecord { Name = blockName };
                bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);
                tr.Commit();
            }
        }

        private static void CreateMarkGroupsTemplateBlock()
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            const string blockName = "MARK_GROUPS"; // имя блока

            using (var lok = doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                //ObjectId brId;
                if (bt.Has(blockName)) return;

                // Создаём пустое определение блока
                bt.UpgradeOpen();
                var btr = new BlockTableRecord { Name = blockName };
                bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);
                tr.Commit();
            }
        }


        private static void DeleteTemplateBlocks()
        {
            var doc = ACAD.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            HashSet<string> blockNames = new HashSet<string> { "SIGN_TEMPLATE", "STAND_TEMPLATE", "MARK_TEMPLATE", "SIGN_GROUPS", "MARK_GROUPS" };

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

        public static void BuildTemplateDWG()
        {
            DeleteTemplateBlocks();
            CreateStandTemplateBlock();
            CreateSignTemplateBlock();
            CreateMarkTemplateBlock();
            CreateSignGroupsTemplateBlock();
            CreateMarkGroupsTemplateBlock();
        }


        //  **************************************************************   ВЫВОД  **************************************************************    //

        // получить список блоков 
        public static List<T> GetListOfRefBlocks<T>(Dictionary<string, string> dictionary = null)
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

                                    // ID
                                    stand.ID = btRefId;

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

                                    // ID
                                    sign.ID = btRefId;

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

                                            case var a when a.Equals("TYPESIZE", StringComparison.OrdinalIgnoreCase):
                                                sign.TypeSize = ar.TextString;
                                                break;

                                            case var a when a.Equals("SIGNEXISTENCE", StringComparison.OrdinalIgnoreCase):
                                                switch (ar.TextString)
                                                {
                                                    case "Установить": sign.Existence = "Необходимо установить"; break;
                                                    case "Демонтировать": sign.Existence = "Необходимо демонтировать"; break;
                                                    default: sign.Existence = ar.TextString; break;
                                                }
                                                break;
                                        }
                                    }

                                    // вытягиваем имя из словаря
                                    if (dictionary != null && dictionary.ContainsKey(sign.Number))
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


                                case Mark mark:

                                    // ID
                                    mark.ID = btRefId;

                                    //Обход атрибутов вставки
                                    foreach (ObjectId arId in brRef.AttributeCollection)
                                    {
                                        var ar = (AttributeReference)tr.GetObject(arId, OpenMode.ForRead);

                                        switch (ar.Tag)
                                        {
                                            case var a when a.Equals("ОСЬ", StringComparison.OrdinalIgnoreCase):
                                                mark.AxisName = ar.TextString;
                                                break;

                                            case var a when a.Equals("MATERIAL", StringComparison.OrdinalIgnoreCase):
                                                mark.Material = ar.TextString;
                                                break;

                                            case var a when a.Equals("NUMBER", StringComparison.OrdinalIgnoreCase):
                                                mark.Number = ar.TextString;
                                                break;

                                            case var a when a.Equals("MARKEXISTENCE", StringComparison.OrdinalIgnoreCase):
                                                switch (ar.TextString)
                                                {
                                                    case "Нанести": mark.Existence = "Требуется нанести"; break;
                                                    case "Демаркировать": mark.Existence = "Требуется демаркировать"; break;
                                                    default: mark.Existence = ar.TextString; break;
                                                }
                                                break;
                                        }
                                    }

                                    // количество
                                    mark.Quantity = "1 шт";

                                    GetMarkProperties(tr, db, mark, brRef);  // Заполняем остальные данные по блоку разметки

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

            Autodesk.AutoCAD.DatabaseServices.Polyline axisPolyline = axis.AxisPoly;   // полилиния оси

            Point3d pkPoint = axisPolyline.GetClosestPointTo(insertPoint, false);   // точка ПК стойки на оси

            // считаем ПК
            double koef = db.Insunits == UnitsValue.Meters ? 1 : 0.001;   // коэффициент перевода в метры
            var distance = axisPolyline.GetDistAtPoint(pkPoint) * koef;   // расстояние от начала 
            if (axis.ReverseDirection) distance = axisPolyline.Length * koef - distance;    // если реверсивное направление оси
            distance = Math.Round(distance, 3) + axis.StartPK * 100;

            stand.Distance = distance;  // расстояние от начала оси до точки ПК

            int pt_1 = (int)Math.Truncate(distance / 100);
            double pt_2 = Math.Round((distance - pt_1 * 100), 2);

            stand.PK = $"ПК {pt_1} + {pt_2}";  // пикет

            // положение слева/справа от оси (на оси)
            stand.Side = GetObjectSide(axis, insertPoint, koef);
        }



        private static void GetMarkProperties(Transaction tr, Database db, Mark mark, BlockReference bref)
        {
            Axis axis = TsoddHost.Current.axis.FirstOrDefault(a => a.Name == mark.AxisName);   // ось к которой привязанна стойка
            if (axis == null)
            {
                MessageBox.Show($"Ошибка. Не получилось найти ось {mark.AxisName} при создании объекта Mark (разметка)." +
                    $" Проверь блок в координатах {bref.Position} ");

                return;
            }

            Autodesk.AutoCAD.DatabaseServices.Polyline axisPolyline = axis.AxisPoly;   // полилиния оси

            double maxPkDistance = double.MinValue;
            double minPkDistance = double.MaxValue;

            List<Point3d> pointsPK = new List<Point3d>();
            Dictionary<double, string> PK = new Dictionary<double, string>();

            // найдем все точки PK_VAL
            double koef = db.Insunits == UnitsValue.Meters ? 1 : 0.001;   // коэффициент перевода в метры
            foreach (ObjectId arId in bref.AttributeCollection)
            {
                var ar = (AttributeReference)tr.GetObject(arId, OpenMode.ForRead);
                if (ar.Tag.Contains("PK_VAL"))
                {
                    Point3d currentPkPosition = ar.Position;
                    pointsPK.Add(currentPkPosition);

                    // считаем ПК
                    double distance = 0;
                    try
                    {
                        distance = axisPolyline.GetDistAtPoint(axisPolyline.GetClosestPointTo(currentPkPosition, false)) * koef;   // расстояние от начала 
                    }
                    catch
                    {
                        MessageBox.Show($"Ошибка определения ПК для блока разметки. Проверь блок в координатах {bref.Position}");
                        continue;
                    }


                    if (axis.ReverseDirection) distance = axisPolyline.Length * koef - distance;    // если реверсивное направление оси
                    distance = Math.Round(distance, 3) + axis.StartPK * 100;

                    if (distance > maxPkDistance) maxPkDistance = distance;
                    if (distance < minPkDistance) minPkDistance = distance;

                    int pt_1 = (int)Math.Truncate(distance / 100);
                    double pt_2 = Math.Round((distance - pt_1 * 100), 2);

                    PK.Add(distance, $"ПК {pt_1} + {pt_2}"); // запоминаем в словарь расстояние и значение ПК
                }
            }

            // считаем площадь разметки
            mark.Square = Math.Round(GetHatchSquare(tr, db, bref), 1);

            // записываем значения начала и конца ПК
            if (PK.TryGetValue(minPkDistance, out string minPkVal)) mark.PK_start = minPkVal;
            if (PK.TryGetValue(maxPkDistance, out string maxPkVal)) mark.PK_end = maxPkVal;

            // считаем среднюю точку и положение относительно оси
            if (pointsPK.Count > 0)
            {
                Point3d averagePoint = new Point3d(pointsPK.Sum(p => p.X) / pointsPK.Count, pointsPK.Sum(p => p.Y) / pointsPK.Count, 0);
                mark.Side = GetObjectSide(axis, averagePoint, koef);
            }
        }


        private static double GetHatchSquare(Transaction tr, Database db, BlockReference bref)
        {
            double sum = 0;
            double koef = db.Insunits == UnitsValue.Meters ? 1 : 0.001;   // коэффициент перевода в метры

            DBObjectCollection explodedObjects = new DBObjectCollection();
            bref.Explode(explodedObjects);

            // перебираем все элементы блока и ищем штриховки
            foreach (DBObject obj in explodedObjects)
            {
                if (obj is BlockReference blr) // если это блок, то рекурсируем
                {
                    sum += GetHatchSquare(tr, db, blr);
                }

                if (obj is Autodesk.AutoCAD.DatabaseServices.Polyline poly) // если это 
                {
                    try { if (poly.ConstantWidth > 0) sum += poly.Length * koef * poly.ConstantWidth * koef; }
                    catch
                    {  
                        try {
                            sum += GetSquareInCompexWidthPoly(poly, koef);
                            }
                        catch{ }
                    }
                }

                if (obj is Hatch hatch)
                {
                    sum += hatch.Area * koef * koef;
                }
            }

            return sum;

            // внутренний метод для полилиний со сложной толщиной (например стрелки)
            double GetSquareInCompexWidthPoly(Polyline poly, double k)
            {
                double result = 0;

                for (int i = 0; i < poly.NumberOfVertices-1; i++)
                {
                    double length = (poly.GetDistanceAtParameter(i + 1) - poly.GetDistanceAtParameter(i))*k;
                    double width_current = poly.GetStartWidthAt(i) * k;
                    double width_next = poly.GetEndWidthAt(i) * k;
                    
                    result += (width_current + width_next) * length / 2 ;
                }

                return result;
            }
        }



        public static List<Mark> GetListOfLineTypes()
        {
            List<Mark> result = new List<Mark>();

            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var objectIds = LineTypeReader.CollectMarkLineTypeID(onlyMaster: true); // получаем ID всех master линий
            if (objectIds.Length == 0) return result; // если нет разметки в виде линий в чертеже, то выходим

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var id in objectIds)
                {
                    Mark currentMarkLineType = new Mark();  // экземпляр разметки
                    TsoddXdataElement xDataElement = new TsoddXdataElement();
                    xDataElement.Parse(id);

                    // площадь разметки основной линии 
                    double koef = db.Insunits == UnitsValue.Meters ? 1 : 0.001;   // коэффициент перевода в метры

                    var masterPolyline = (Autodesk.AutoCAD.DatabaseServices.Polyline)tr.GetObject(xDataElement.MasterPolylineID, OpenMode.ForRead);  // получаем полилинию по ID

                    // полилиния оси
                    Handle tempHandle = new Handle(Convert.ToInt64(xDataElement.AxisHandle, 16));
                    ObjectId axisID = db.GetObjectId(false, tempHandle, 0);

                    Axis axis = TsoddHost.Current.axis.FirstOrDefault(a => a.PolyID == axisID);   // ось к которой привязанна линия

                    currentMarkLineType.AxisName = axis.Name;       // имя привязанной оси

                    currentMarkLineType.Number = xDataElement.Number;       // номер разметки

                    currentMarkLineType.Quantity = $"{Math.Round(masterPolyline.Length * koef, 1)} м";      // количество

                    currentMarkLineType.Square = GetPatternValues(masterPolyline, koef); // площадь master полилинии

                    if (xDataElement.SlavePolylineID != ObjectId.Null)      // если десть slave полилиния
                    {
                        var slavePolyline = (Autodesk.AutoCAD.DatabaseServices.Polyline)tr.GetObject(xDataElement.SlavePolylineID, OpenMode.ForRead);  // получаем полилинию по ID
                        currentMarkLineType.Square += GetPatternValues(slavePolyline, koef); // площадь slave полилинии
                    }

                    currentMarkLineType.Square = Math.Round(currentMarkLineType.Square, 1);

                    double dist_1 = 0;  // переменные для опрделения расстояний ПК
                    double dist_2 = 0;

                    try
                    {
                        dist_1 = axis.AxisPoly.GetDistAtPoint(axis.AxisPoly.GetClosestPointTo(masterPolyline.StartPoint, false)) * koef;
                        dist_2 = axis.AxisPoly.GetDistAtPoint(axis.AxisPoly.GetClosestPointTo(masterPolyline.EndPoint, false)) * koef;
                    }
                    catch
                    {
                        MessageBox.Show($"Ошибка определения ПК для  линии разметки. Проверь линию с началом в координатах {masterPolyline.StartPoint}");
                        continue;
                    }

                    if (axis.ReverseDirection) // если реверсивное направление оси
                    {
                        dist_1 = Math.Round(axis.AxisPoly.Length * koef - dist_1, 3) + axis.StartPK * 100;
                        dist_2 = Math.Round(axis.AxisPoly.Length * koef - dist_2, 3) + axis.StartPK * 100;
                    }

                    double minDist = Math.Min(dist_1, dist_2);
                    double maxDist = Math.Max(dist_1, dist_2);

                    int pt_1 = 0;
                    double pt_2 = 0;

                    pt_1 = (int)Math.Truncate(minDist / 100);
                    pt_2 = Math.Round((minDist - pt_1 * 100), 2);
                    currentMarkLineType.Distance = minDist;
                    currentMarkLineType.PK_start = $"ПК {pt_1} + {pt_2}";   // ПК начала линии

                    pt_1 = (int)Math.Truncate(maxDist / 100);
                    pt_2 = Math.Round((maxDist - pt_1 * 100), 2);
                    currentMarkLineType.PK_end = $"ПК {pt_1} + {pt_2}";     // ПК конца линии

                    currentMarkLineType.Side = GetObjectSide(axis, masterPolyline.StartPoint, koef);    // сторона относительно оси

                    currentMarkLineType.Material = xDataElement.Material;

                    switch (xDataElement.Existence)
                    {
                        case "Нанести": currentMarkLineType.Existence = "Требуется нанести"; break;
                        case "Демаркировать": currentMarkLineType.Existence = "Требуется демаркировать"; break;
                        default: currentMarkLineType.Existence = xDataElement.Existence; break;
                    }

                    result.Add(currentMarkLineType);
                }
            }

            // метод для определения площади линии
            double GetPatternValues(Autodesk.AutoCAD.DatabaseServices.Polyline polyline, double koef)
            {
                (double dash, double space) patternValues = (new double(), new double());
                string lineTypeName = polyline.Linetype;

                // если сплошная линия
                if (lineTypeName.Contains("continuous"))
                {
                    return polyline.Length * polyline.ConstantWidth * Math.Pow(koef, 2);
                }

                int start = lineTypeName.IndexOf('(') + 1;
                int end = lineTypeName.IndexOf(')', start);
                var values = lineTypeName.Substring(start, end - start).Split('_');
                patternValues.dash = values.Select(x => double.Parse(x)).Where(x => x > 0).Sum();
                patternValues.space = Math.Abs(values.Select(x => double.Parse(x)).Where(x => x < 0).Sum());

                double dash_k = patternValues.dash / (patternValues.dash + patternValues.space);
                double dash_s = polyline.Length * dash_k * polyline.ConstantWidth * Math.Pow(koef, 2);

                return dash_s;
            }
            return result;
        }


        // метод определяет сторону положения объекта относительно оси
        private static string GetObjectSide(Axis axis, Point3d objPoint, double unitsKoef)
        {
            Point3d pkPoint = axis.AxisPoly.GetClosestPointTo(objPoint, false);   // точка объекта, спроецированная на ось на оси
            var minDist = 1 / unitsKoef;     // минимальное расстояние, которое определяет положение объекта " на оси"

            double distToAxis = objPoint.DistanceTo(pkPoint);   // расстояние от объекта до оси

            if (distToAxis <= minDist)
            {
                return "На оси";
            }
            else
            {
                try
                {
                    // расстояние бдля определения вектора по оси с учетом направления оси
                    double vectDist = axis.ReverseDirection ? axis.AxisPoly.GetDistAtPoint(pkPoint) + 1 :
                                                              axis.AxisPoly.GetDistAtPoint(pkPoint) - 1;
                    // точка для вектора на оси
                    Point3d secondPointOnAxis = axis.AxisPoly.GetPointAtDist(vectDist);

                    Vector3d projectionVector = (objPoint - pkPoint).GetNormal();   // вектор, направленный по проекции объекта на ось
                    Point3d projectionPoint = pkPoint + projectionVector * 1;       // точка на проекции объекта на ось


                    Vector3d vectorAxis = (pkPoint - secondPointOnAxis).GetNormal();
                    Vector3d vrctorPoint = (projectionPoint - secondPointOnAxis).GetNormal();

                    var cross = vectorAxis.CrossProduct(vrctorPoint);
                    return cross.Z < 0 ? "Справа от оси" : "Слева от оси";
                }
                catch
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage
                        ($"Не удалось отсроить вектор для линии разметки с началом в координатах {objPoint}");

                    return "Не получилось определить";
                }
            }
        }


    }
}











