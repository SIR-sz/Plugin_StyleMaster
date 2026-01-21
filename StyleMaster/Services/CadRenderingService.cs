using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using StyleMaster.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace StyleMaster.Services
{
    public static class CadRenderingService
    {
        /// <summary>
        /// 全量同步填充：先清除旧填充，再根据列表重新生成并强制重绘。
        /// </summary>
        public static void ExecuteFill(IEnumerable<MaterialItem> items)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var targetLayerNames = items.Select(x => x.LayerName).ToList();

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    // 1. 清理
                    ClearInternal(tr, btr, targetLayerNames);

                    // 2. 生成
                    foreach (var item in items)
                    {
                        var boundaryIds = GetEntitiesOnLayer(btr, tr, item.LayerName);
                        foreach (ObjectId bId in boundaryIds)
                        {
                            if (item.FillMode == FillType.Hatch)
                                CreateHatch(tr, btr, bId, item);
                        }
                    }

                    // 3. 排序
                    ReorderDrawOrder(tr, btr, items);
                    tr.Commit();
                }

                // ✨ 深度重绘逻辑：冲洗图形队列并重生成
                doc.TransactionManager.QueueForGraphicsFlush();
                doc.Editor.Regen();
            }
        }

        /// <summary>
        /// 局部刷新单层：解决比例修改不生效的核心入口。
        /// </summary>
        public static void RefreshSingleLayer(MaterialItem item, IEnumerable<MaterialItem> allItems)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (doc.LockDocument())
            {
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var btr = (BlockTableRecord)tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);

                    // 清理单层
                    ClearInternal(tr, btr, new List<string> { item.LayerName });

                    // 生成新填充
                    var boundaryIds = GetEntitiesOnLayer(btr, tr, item.LayerName);
                    foreach (ObjectId bId in boundaryIds)
                    {
                        if (item.FillMode == FillType.Hatch)
                            CreateHatch(tr, btr, bId, item);
                    }

                    ReorderDrawOrder(tr, btr, allItems);
                    tr.Commit();
                }

                // ✨ 深度重绘逻辑
                doc.TransactionManager.QueueForGraphicsFlush();
                doc.Editor.Regen();
            }
        }

        /// <summary>
        /// 清理指定图层的填充。
        /// </summary>
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

        private static void ClearInternal(Transaction tr, BlockTableRecord btr, List<string> layers)
        {
            foreach (ObjectId id in btr)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent is Hatch hatch && layers.Any(n => n.Equals(hatch.Layer, StringComparison.OrdinalIgnoreCase)))
                {
                    tr.GetObject(id, OpenMode.ForWrite);
                    ent.Erase();
                }
            }
        }

        private static void CreateHatch(Transaction tr, BlockTableRecord btr, ObjectId boundaryId, MaterialItem settings)
        {
            Hatch hat = new Hatch();
            hat.SetDatabaseDefaults();
            hat.Layer = settings.LayerName;

            btr.AppendEntity(hat);
            tr.AddNewlyCreatedDBObject(hat, true);

            hat.Associative = true;
            hat.Normal = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis;
            hat.Color = settings.CadColor;

            // ✨ 核心修复步骤：
            // 1. 设置图案类型和名称
            hat.SetHatchPattern(HatchPatternType.PreDefined, settings.PatternName);

            // 2. 设置比例数值
            hat.PatternScale = settings.Scale;

            // 3. 【关键魔术】再次设置一遍图案！
            // 这一步会强制 Hatch 引擎根据上面刚设好的 PatternScale 重新计算图案线
            hat.SetHatchPattern(hat.PatternType, hat.PatternName);

            hat.Transparency = new Transparency((byte)(255 * (1 - settings.Opacity / 100.0)));
            hat.AppendLoop(HatchLoopTypes.Outermost, new ObjectIdCollection { boundaryId });

            // 4. 评估并记录图形修改
            hat.EvaluateHatch(true);
            hat.RecordGraphicsModified(true);
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
                    if (ent is Hatch h && h.Layer.Equals(item.LayerName, StringComparison.OrdinalIgnoreCase))
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
                if (ent != null && ent.Layer.Equals(layerName, StringComparison.OrdinalIgnoreCase) && !(ent is Hatch))
                    ids.Add(id);
            }
            return ids;
        }
    }
}