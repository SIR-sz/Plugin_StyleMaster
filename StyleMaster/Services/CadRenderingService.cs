using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using StyleMaster.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace StyleMaster.Services
{
    public static class CadRenderingService
    {
        public static void ExecuteFill(IEnumerable<MaterialItem> items)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            ed.WriteMessage("\n[StyleMaster] 开始执行全量同步...");

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    var targetLayers = items.Select(x => x.LayerName).ToList();

                    ClearInternal(tr, btr, targetLayers);

                    foreach (var item in items)
                    {
                        var boundaryIds = GetEntitiesOnLayer(btr, tr, item.LayerName);
                        ed.WriteMessage($"\n[图层: {item.LayerName}] 找到 {boundaryIds.Count} 个边界对象。");

                        foreach (ObjectId bId in boundaryIds)
                        {
                            if (item.FillMode == StyleMaster.Models.FillType.Hatch)
                                CreateHatch(tr, btr, bId, item, ed);
                            else if (item.FillMode == StyleMaster.Models.FillType.Image)
                                CreateImageFill(tr, btr, bId, item, ed);
                        }
                    }

                    ReorderDrawOrder(tr, btr, items);
                    tr.Commit();
                }
                ed.Regen();
            }
            ed.WriteMessage("\n[StyleMaster] 同步完成。");
        }

        // 局部刷新增加调试
        public static void RefreshSingleLayer(MaterialItem item, IEnumerable<MaterialItem> allItems)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            ed.WriteMessage($"\n[StyleMaster] 正在刷新图层: {item.LayerName}");

            using (doc.LockDocument())
            {
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var btr = (BlockTableRecord)tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);
                    ClearInternal(tr, btr, new List<string> { item.LayerName });

                    var ids = GetEntitiesOnLayer(btr, tr, item.LayerName);
                    foreach (ObjectId bId in ids)
                    {
                        if (item.FillMode == StyleMaster.Models.FillType.Hatch)
                            CreateHatch(tr, btr, bId, item, ed);
                        else if (item.FillMode == StyleMaster.Models.FillType.Image)
                            CreateImageFill(tr, btr, bId, item, ed);
                    }

                    ReorderDrawOrder(tr, btr, allItems);
                    tr.Commit();
                }
                ed.Regen();
            }
        }

        // 核心图片填充：带每步调试
        private static void CreateImageFill(Transaction tr, BlockTableRecord btr, ObjectId boundaryId, MaterialItem settings, Editor ed)
        {
            // 1. 基础路径检查
            string assemblyDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string imagePath = Path.Combine(assemblyDir, "Resources", "Materials", settings.PatternName);
            if (!File.Exists(imagePath)) return;

            // 2. 确保图片定义存在 (ImageDef)
            ObjectId imageDefId = GetOrCreateImageDef(btr.Database, tr, imagePath);
            if (imageDefId.IsNull) return;

            // 3. 获取边界
            var pline = tr.GetObject(boundaryId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Polyline;
            if (pline == null) return;
            Extents3d ext = pline.GeometricExtents;

            // 4. 创建图片实体
            RasterImage image = new RasterImage();
            image.SetDatabaseDefaults();
            image.ImageDefId = imageDefId;
            image.Layer = settings.LayerName;

            double w = ext.MaxPoint.X - ext.MinPoint.X;
            double h = ext.MaxPoint.Y - ext.MinPoint.Y;
            image.Orientation = new CoordinateSystem3d(ext.MinPoint, new Vector3d(w, 0, 0), new Vector3d(0, h, 0));
            image.ImageTransparency = true;

            // 将图片添加到数据库
            ObjectId imageId = btr.AppendEntity(image);
            tr.AddNewlyCreatedDBObject(image, true);

            // ✨ 核心修复：只记录句柄，不在这里执行命令
            string imgH = imageId.Handle.ToString();
            string plH = boundaryId.Handle.ToString();

            // 构建一段纯净的 LISP 脚本
            // 使用 vl-catch-all-apply 确保即使一个填充失败，也不会弹出报错对话框
            string lispCmd = string.Format(
                "(vl-catch-all-apply '(lambda () " +
                "  (setvar \"CMDECHO\" 0) " +
                "  (command \"_.IMAGECLIP\" (handent \"{0}\") \"_N\" \"_S\" (handent \"{1}\")) " +
                "  (setvar \"IMAGEFRAME\" 0) " +
                "))(princ) ",
                imgH, plH
            );

            // ✨ 核心修复：通过异步队列发送
            // 注意：不要在之后访问 'image' 对象的任何属性，因为它可能已被事务回收
            Document doc = Application.DocumentManager.MdiActiveDocument;
            doc.SendStringToExecute(lispCmd, true, false, false);
        }

        // ✨ 辅助函数：计算多边形面积以判断方向
        private static double GetArea(List<Point2d> pts)
        {
            double area = 0;
            for (int i = 0; i < pts.Count; i++)
            {
                int j = (i + 1) % pts.Count;
                area += pts[i].X * pts[j].Y;
                area -= pts[j].X * pts[i].Y;
            }
            return area / 2.0;
        }

        // 填充调试
        private static void CreateHatch(Transaction tr, BlockTableRecord btr, ObjectId boundaryId, MaterialItem settings, Editor ed)
        {
            try
            {
                string patternName = string.IsNullOrEmpty(settings.PatternName) ? "ANSI31" : settings.PatternName;

                Hatch hat = new Hatch();
                hat.SetDatabaseDefaults();
                hat.Layer = settings.LayerName;
                btr.AppendEntity(hat);
                tr.AddNewlyCreatedDBObject(hat, true);

                hat.Associative = true;
                hat.Normal = Vector3d.ZAxis;
                hat.Color = settings.CadColor;

                try
                {
                    hat.SetHatchPattern(HatchPatternType.PreDefined, patternName);
                    hat.PatternScale = settings.Scale > 0 ? settings.Scale : 1.0;
                    hat.SetHatchPattern(hat.PatternType, hat.PatternName);
                }
                catch
                {
                    ed.WriteMessage($"\n  [警告] 图案 {patternName} 无效，退回到 SOLID");
                    hat.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                }

                hat.Transparency = new Transparency((byte)(255 * (1 - settings.Opacity / 100.0)));
                hat.AppendLoop(HatchLoopTypes.Outermost, new ObjectIdCollection { boundaryId });
                hat.EvaluateHatch(true);
                hat.RecordGraphicsModified(true);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n  [错误] 填充失败: {ex.Message}");
            }
        }

        // 其他辅助方法保持不变 (GetEntitiesOnLayer, ClearInternal, ReorderDrawOrder, GetOrCreateImageDef)
        // ... 请务必保留之前版本中这些方法的实现 ...

        public static void ClearFillsOnLayers(IEnumerable<string> layerNames)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            using (doc.LockDocument())
            {
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var btr = (BlockTableRecord)tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);
                    ClearInternal(tr, btr, layerNames.ToList());
                    tr.Commit();
                }
                doc.Editor.Regen();
            }
        }

        private static ObjectId GetOrCreateImageDef(Database db, Transaction tr, string imagePath)
        {
            ObjectId dictId = RasterImageDef.GetImageDictionary(db);
            if (dictId.IsNull) dictId = RasterImageDef.CreateImageDictionary(db);
            DBDictionary dict = tr.GetObject(dictId, OpenMode.ForWrite) as DBDictionary;
            string fileName = Path.GetFileNameWithoutExtension(imagePath);
            if (dict.Contains(fileName)) return dict.GetAt(fileName);
            RasterImageDef imgDef = new RasterImageDef();
            imgDef.SourceFileName = imagePath;
            imgDef.Load();
            ObjectId id = dict.SetAt(fileName, imgDef);
            tr.AddNewlyCreatedDBObject(imgDef, true);
            return id;
        }

        private static void ClearInternal(Transaction tr, BlockTableRecord btr, List<string> layers)
        {
            foreach (ObjectId id in btr)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if ((ent is Hatch || ent is RasterImage) && layers.Any(n => n.Equals(ent.Layer, StringComparison.OrdinalIgnoreCase)))
                {
                    tr.GetObject(id, OpenMode.ForWrite);
                    ent.Erase();
                }
            }
        }

        private static void ReorderDrawOrder(Transaction tr, BlockTableRecord btr, IEnumerable<MaterialItem> items)
        {
            var dotId = btr.DrawOrderTableId;
            if (dotId.IsNull) return;
            var dot = (DrawOrderTable)tr.GetObject(dotId, OpenMode.ForWrite);
            var sorted = items.OrderByDescending(x => x.Priority).ToList();
            foreach (var item in sorted)
            {
                ObjectIdCollection ids = new ObjectIdCollection();
                foreach (ObjectId id in btr)
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if ((ent is Hatch || ent is RasterImage) && ent.Layer.Equals(item.LayerName, StringComparison.OrdinalIgnoreCase))
                        ids.Add(id);
                }
                if (ids.Count > 0) dot.MoveToTop(ids);
            }
        }

        private static ObjectIdCollection GetEntitiesOnLayer(BlockTableRecord btr, Transaction tr, string layerName)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            foreach (ObjectId id in btr)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent != null && ent.Layer.Equals(layerName, StringComparison.OrdinalIgnoreCase) && !(ent is Hatch) && !(ent is RasterImage))
                    ids.Add(id);
            }
            return ids;
        }
    }
}