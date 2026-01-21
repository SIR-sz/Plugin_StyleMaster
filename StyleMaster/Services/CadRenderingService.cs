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
        /// 全量同步填充：先清除旧填充，再根据列表重新生成。
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
                    foreach (ObjectId id in btr)
                    {
                        var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent is Hatch hatch && targetLayerNames.Any(n => n.Equals(hatch.Layer, StringComparison.OrdinalIgnoreCase)))
                        {
                            tr.GetObject(id, OpenMode.ForWrite);
                            ent.Erase();
                        }
                    }

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
                doc.Editor.UpdateScreen();
            }
        }

        /// <summary>
        /// 局部刷新单层：删除该层旧填充并重新填充，随后修正层级。
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

                    foreach (ObjectId id in btr)
                    {
                        var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent is Hatch hatch && hatch.Layer.Equals(item.LayerName, StringComparison.OrdinalIgnoreCase))
                        {
                            tr.GetObject(id, OpenMode.ForWrite);
                            ent.Erase();
                        }
                    }

                    var boundaryIds = GetEntitiesOnLayer(btr, tr, item.LayerName);
                    foreach (ObjectId bId in boundaryIds)
                    {
                        if (item.FillMode == FillType.Hatch)
                            CreateHatch(tr, btr, bId, item);
                    }

                    ReorderDrawOrder(tr, btr, allItems);
                    tr.Commit();
                }
                doc.Editor.UpdateScreen();
            }
        }

        /// <summary>
        /// ✨ 修复：补充缺失的清除图层填充方法
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
                    int count = 0;

                    foreach (ObjectId id in btr)
                    {
                        var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent is Hatch hatch && layerNames.Any(l => l.Equals(hatch.Layer, StringComparison.OrdinalIgnoreCase)))
                        {
                            tr.GetObject(id, OpenMode.ForWrite);
                            ent.Erase();
                            count++;
                        }
                    }
                    tr.Commit();
                    doc.Editor.WriteMessage($"\n[StyleMaster] 已成功清理 {count} 个填充对象。");
                }
                doc.Editor.UpdateScreen();
            }
        }

        private static void CreateHatch(Transaction tr, BlockTableRecord btr, ObjectId boundaryId, MaterialItem settings)
        {
            Hatch hat = new Hatch();
            hat.SetDatabaseDefaults();
            hat.Layer = settings.LayerName;
            btr.AppendEntity(hat);
            tr.AddNewlyCreatedDBObject(hat, true);

            hat.Normal = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis;
            hat.Color = settings.CadColor;
            hat.SetHatchPattern(HatchPatternType.PreDefined, settings.PatternName);
            hat.PatternScale = settings.Scale;
            hat.Transparency = new Transparency((byte)(255 * (1 - settings.Opacity / 100.0)));
            hat.AppendLoop(HatchLoopTypes.Outermost, new ObjectIdCollection { boundaryId });
            hat.EvaluateHatch(true);
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