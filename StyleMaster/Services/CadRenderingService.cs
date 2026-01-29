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
using System.Text;

namespace StyleMaster.Services
{
    public static class CadRenderingService
    {
        /// <summary>
        /// 将指定的材质项对应图层的边界对象导出为 SVG 文件，按图层使用 <g> 标签分组。
        /// </summary>
        public static void ExportToSvg(IEnumerable<MaterialItem> items, string savePath)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;

            try
            {
                using (doc.LockDocument())
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                        StringBuilder svg = new StringBuilder();
                        svg.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                        // viewBox 建议根据图纸实际范围动态计算，此处暂设固定大值
                        svg.AppendLine("<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 10000 10000\">");

                        foreach (var item in items)
                        {
                            var boundaryIds = GetEntitiesOnLayer(btr, tr, item.LayerName);
                            if (boundaryIds.Count == 0) continue;

                            // 用 <g> 标签区分 CAD 图层，并带上 data-pattern 供 PS 解析
                            svg.AppendLine($"  <g id=\"{item.LayerName}\" data-pattern=\"{item.PatternName}\" fill=\"none\" stroke=\"black\" stroke-width=\"1\">");

                            foreach (ObjectId id in boundaryIds)
                            {
                                if (tr.GetObject(id, OpenMode.ForRead) is Polyline pl)
                                {
                                    svg.Append("    <path d=\"M ");
                                    for (int i = 0; i < pl.NumberOfVertices; i++)
                                    {
                                        Point2d pt = pl.GetPoint2dAt(i);
                                        // CAD Y 轴向上，SVG Y 轴向下，需取反处理
                                        svg.Append($"{pt.X} {-pt.Y} ");
                                        if (i < pl.NumberOfVertices - 1) svg.Append("L ");
                                    }
                                    svg.AppendLine("Z\" />");
                                }
                            }
                            svg.AppendLine("  </g>");
                        }
                        svg.AppendLine("</svg>");

                        File.WriteAllText(savePath, svg.ToString());
                        tr.Commit();
                    }
                }
                doc.Editor.WriteMessage($"\n[StyleMaster] SVG 已成功导出至: {savePath}");
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage($"\n[错误] 导出失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 核心执行逻辑：在 CAD 内仅执行 Hatch 填充，确保轻量化。
        /// </summary>
        public static void ExecuteFill(IEnumerable<MaterialItem> items)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            ed.WriteMessage("\n[StyleMaster] 正在执行 Hatch 同步...");

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
                        foreach (ObjectId bId in boundaryIds)
                        {
                            CreateHatch(tr, btr, bId, item, ed);
                        }
                    }

                    ReorderDrawOrder(tr, btr, items);
                    tr.Commit();
                }
                ed.Regen();
            }
            ed.WriteMessage("\n[StyleMaster] CAD 填充完成。");
        }

        /// <summary>
        /// 刷新单个图层的填充。
        /// </summary>
        public static void RefreshSingleLayer(MaterialItem item, IEnumerable<MaterialItem> allItems)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            using (doc.LockDocument())
            {
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var btr = (BlockTableRecord)tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);
                    ClearInternal(tr, btr, new List<string> { item.LayerName });

                    var ids = GetEntitiesOnLayer(btr, tr, item.LayerName);
                    foreach (ObjectId bId in ids)
                    {
                        CreateHatch(tr, btr, bId, item, ed);
                    }

                    ReorderDrawOrder(tr, btr, allItems);
                    tr.Commit();
                }
                ed.Regen();
            }
        }

        /// <summary>
        /// 创建 Hatch 实体。
        /// </summary>
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
                    hat.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                }

                hat.Transparency = new Transparency((byte)(255 * (1 - settings.Opacity / 100.0)));
                hat.AppendLoop(HatchLoopTypes.Outermost, new ObjectIdCollection { boundaryId });
                hat.EvaluateHatch(true);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n  [错误] 填充失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 内部清理逻辑：清理指定图层上的所有 Hatch 和残留的 RasterImage 对象。
        /// </summary>
        private static void ClearInternal(Transaction tr, BlockTableRecord btr, List<string> layers)
        {
            foreach (ObjectId id in btr)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                if ((ent is Hatch || ent is RasterImage) &&
                    layers.Any(n => n.Equals(ent.Layer, StringComparison.OrdinalIgnoreCase)))
                {
                    tr.GetObject(id, OpenMode.ForWrite);
                    ent.Erase();
                }
            }
        }

        /// <summary>
        /// 调整图层显示顺序。
        /// </summary>
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

        /// <summary>
        /// 获取指定图层上的边界实体（排除填充和图片）。
        /// </summary>
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
    }
}