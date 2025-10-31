using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(ContOnPlugin.Commands))]

namespace ContOnPlugin
{
    public class Commands
    {
        // Настройки
        const string TargetBlockName = "GLB_CONT_L";
        const string TagEht = "EHT_CONT";
        const string TagLine = "LINE_NUMBER";
        const string TagEast = "EAST_COORD";
        const string TagNorth = "NORTH_COORD";
        const string TagElev = "ELEV";
        const string TagExtra = "PIPING_CONT_X"; // скрытый/выдвижной
        const double XTol = 5.0;                  // допуск по X (чертёжные единицы)

        // Жёстко заданная папка с символами (по вашему пути)
        static readonly string SymbolsFolder = @"C:\Users\E2022157\nVent Management Company\EED CAD - IHTS_CAD_Tools\Support_Menus\EMEA\EMEA_Symbols";
        static readonly string DefaultBlockDwg = @"C:\Users\E2022157\nVent Management Company\EED CAD - IHTS_CAD_Tools\Support_Menus\EMEA\EMEA_Symbols\GLB_CONT_L.dwg";

        [CommandMethod("CONT2BLK")]
        public void ConvertStacksToBlock()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Текущее пространство
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var btrId = db.TileMode == true ? bt[BlockTableRecord.ModelSpace] : bt[BlockTableRecord.PaperSpace];
                var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                // Собираем все MTEXT
                var allMtexts = new List<MTextInfo>();
                foreach (ObjectId id in btr)
                {
                    if (!id.ObjectClass.DxfName.Equals("MTEXT", StringComparison.OrdinalIgnoreCase)) continue;
                    var mt = (MText)tr.GetObject(id, OpenMode.ForRead);
                    allMtexts.Add(new MTextInfo { Id = id, Y = mt.Location.Y, X = mt.Location.X, Text = Clean(mt.Contents), Height = mt.TextHeight });
                }

                if (allMtexts.Count == 0)
                {
                    ed.WriteMessage("\\nMTEXT не найден.");
                    return;
                }

                // Проверяем наличие блока, иначе попробуем загрузить из заданной папки
                EnsureBlockExists(db, tr, ed);

                int replaced = 0, skipped = 0;

                // Ищем семена: строки точно равные CONT. ON
                foreach (var seed in allMtexts.Where(m => string.Equals(m.Text, "CONT. ON", StringComparison.OrdinalIgnoreCase)))
                {
                    // Собираем стопку: те же X (±XTol) и ниже по Y
                    var stack = allMtexts
                        .Where(m => Math.Abs(m.X - seed.X) <= XTol && m.Y <= seed.Y + seed.Height * 6)
                        .OrderByDescending(m => m.Y) // сверху вниз
                        .ToList();

                    // Нормализуем на случай посторонних строк — оставим подряд идущие от seed вниз
                    int seedIndex = stack.FindIndex(m => m.Id == seed.Id);
                    if (seedIndex < 0) { skipped++; continue; }
                    stack = stack.Skip(seedIndex).ToList();

                    // Разбор
                    string lineNo = stack.ElementAtOrDefault(1)?.Text;
                    string extra = null, east = null, north = null, elev = null;
                    var rest = stack.Skip(2).Select(s => s.Text).ToList();

                    if (rest.Count > 0 && !StartsWith(rest[0], "E"))
                    {
                        extra = rest[0];
                        rest = rest.Skip(1).ToList();
                    }

                    foreach (var s in rest)
                    {
                        if (east == null && StartsWith(s, "E")) east = s;
                        else if (north == null && StartsWith(s, "N")) north = s;
                        else if (elev == null && StartsWith(s, "EL")) elev = s;
                    }

                    if (string.IsNullOrWhiteSpace(lineNo) || east == null || north == null || elev == null)
                    {
                        skipped++;
                        continue;
                    }

                    // Вставка блока в точке seed
                    btr.UpgradeOpen();
                    var br = new BlockReference(new Point3d(seed.X, seed.Y, 0), bt[TargetBlockName]);
                    btr.AppendEntity(br); tr.AddNewlyCreatedDBObject(br, true);

                    // Создаём атрибуты из определения
                    var bdef = (BlockTableRecord)tr.GetObject(bt[TargetBlockName], OpenMode.ForRead);
                    foreach (ObjectId attId in bdef)
                    {
                        if (!attId.ObjectClass.DxfName.Equals("ATTDEF", StringComparison.OrdinalIgnoreCase)) continue;
                        var ad = (AttributeDefinition)tr.GetObject(attId, OpenMode.ForRead);
                        var ar = new AttributeReference();
                        ar.SetAttributeFromBlock(ad, br.BlockTransform);
                        ar.TextString = ad.TextString;
                        ar.Tag = ad.Tag;

                        // Заполнение по тегу
                        switch (ad.Tag.ToUpperInvariant())
                        {
                            case TagEht:   ar.TextString = "CONT. ON"; break;
                            case TagLine:  ar.TextString = lineNo; break;
                            case TagEast:  ar.TextString = east; break;
                            case TagNorth: ar.TextString = north; break;
                            case TagElev:  ar.TextString = elev; break;
                            case TagExtra:
                                if (!string.IsNullOrWhiteSpace(extra))
                                {
                                    ar.TextString = extra; ar.Invisible = false;
                                }
                                else
                                {
                                    ar.TextString = string.Empty; ar.Invisible = true;
                                }
                                break;
                        }

                        br.AttributeCollection.AppendAttribute(ar);
                        tr.AddNewlyCreatedDBObject(ar, true);
                    }

                    // Удаляем исходные MTEXT этой стопки
                    foreach (var s in stack)
                    {
                        var ent = (Entity)tr.GetObject(s.Id, OpenMode.ForWrite);
                        ent.Erase();
                    }

                    replaced++;
                }

                tr.Commit();
                ed.WriteMessage($@"\nГотово. Заменено: {replaced}, пропущено: {skipped}.");
            }
        }

        private static void EnsureBlockExists(Database db, Transaction tr, Editor ed)
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            if (bt.Has(TargetBlockName)) return;

            // Сначала попытаемся взять GLB_CONT_L.dwg из жёсткого пути
            if (File.Exists(DefaultBlockDwg))
            {
                using (var xdb = new Database(false, true))
                {
                    xdb.ReadDwgFile(DefaultBlockDwg, FileShare.Read, true, "");
                    db.Insert(TargetBlockName, xdb, false);
                    return;
                }
            }

            // Если не нашли, попробуем найти DWG в папке SymbolsFolder
            if (Directory.Exists(SymbolsFolder))
            {
                foreach (var dwg in Directory.EnumerateFiles(SymbolsFolder, "*.dwg"))
                {
                    try
                    {
                        using (var xdb = new Database(false, true))
                        {
                            xdb.ReadDwgFile(dwg, FileShare.Read, true, "");
                            using (var tr2 = xdb.TransactionManager.StartTransaction())
                            {
                                var bt2 = (BlockTable)tr2.GetObject(xdb.BlockTableId, OpenMode.ForRead);
                                if (bt2.Has(TargetBlockName))
                                {
                                    db.Insert(TargetBlockName, xdb, false);
                                    return;
                                }
                            }
                        }
                    }
                    catch {}
                }
            }

            // Фоллбэк — спросить файл у пользователя
            var opt = new PromptOpenFileOptions("Укажите DWG с определением блока '" + TargetBlockName + "'")
            {
                Filter = "Drawing (*.dwg)|*.dwg|All files (*.*)|*.*",
                InitialDirectory = SymbolsFolder
            };
            var res = ed.GetFileNameForOpen(opt);
            if (res.Status != PromptStatus.OK) throw new System.Exception("Блок не найден и файл не выбран.");

            using (var xdb2 = new Database(false, true))
            {
                xdb2.ReadDwgFile(res.StringResult, FileShare.Read, true, "");
                db.Insert(TargetBlockName, xdb2, false);
            }
        }

        private static bool StartsWith(string s, string prefix)
        { return !string.IsNullOrWhiteSpace(s) && s.TrimStart().ToUpperInvariant().StartsWith(prefix.ToUpperInvariant()); }

        private static string Clean(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            return s.Replace("\\P", " ").Replace("\\A1;", string.Empty).Replace("\\A0;", string.Empty).Trim();
        }

        private class MTextInfo
        {
            public ObjectId Id; public double X; public double Y; public string Text; public double Height;
        }
    }
}