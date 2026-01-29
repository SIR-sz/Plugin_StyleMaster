using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
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
        /// 自动化打印表现层线稿为 PDF。
        /// 修改：修正了 PlotSettings 属性引用和打印引擎初始化逻辑。
        /// </summary>
        public static void PlotRepresentationPdf(Autodesk.AutoCAD.DatabaseServices.Extents3d ext, string pdfPath, System.Collections.Generic.List<string> fillLayerNames)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // 1. 隐藏填充层
                    var lt = (Autodesk.AutoCAD.DatabaseServices.LayerTable)tr.GetObject(db.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    var layersToRestore = new System.Collections.Generic.List<Autodesk.AutoCAD.DatabaseServices.LayerTableRecord>();
                    foreach (Autodesk.AutoCAD.DatabaseServices.ObjectId id in lt)
                    {
                        var ltr = (Autodesk.AutoCAD.DatabaseServices.LayerTableRecord)tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        if (fillLayerNames.Contains(ltr.Name) && !ltr.IsOff)
                        {
                            ltr.IsOff = true; // 临时关闭填充层
                            layersToRestore.Add(ltr);
                        }
                    }

                    // 2. 配置打印参数
                    var lo = (Autodesk.AutoCAD.DatabaseServices.Layout)tr.GetObject(Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current.GetLayoutId(Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current.CurrentLayout), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    Autodesk.AutoCAD.PlottingServices.PlotInfo pi = new Autodesk.AutoCAD.PlottingServices.PlotInfo { Layout = lo.ObjectId };

                    Autodesk.AutoCAD.DatabaseServices.PlotSettings ps = new Autodesk.AutoCAD.DatabaseServices.PlotSettings(lo.ModelType);
                    ps.CopyFrom(lo);
                    Autodesk.AutoCAD.DatabaseServices.PlotSettingsValidator psv = Autodesk.AutoCAD.DatabaseServices.PlotSettingsValidator.Current;

                    // 设置打印窗口范围
                    Autodesk.AutoCAD.DatabaseServices.Extents2d windowExt = new Autodesk.AutoCAD.DatabaseServices.Extents2d(ext.MinPoint.X, ext.MinPoint.Y, ext.MaxPoint.X, ext.MaxPoint.Y);
                    psv.SetPlotWindowArea(ps, windowExt);
                    psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
                    psv.SetUseStandardScale(ps, true);
                    psv.SetStdScaleType(ps, Autodesk.AutoCAD.DatabaseServices.StdScaleType.ScaleToFit);
                    psv.SetPlotCentered(ps, true);

                    // 设置 PDF 打印机
                    psv.SetPlotConfigurationName(ps, "AutoCAD PDF (High Quality Print).pc3", "ISO_full_bleed_A3_(420.00_x_297.00_MM)");

                    // ✨ 修改：修正 StyleSheet 获取路径 (从 ps.CurrentStyleSheet 获取)
                    if (!string.IsNullOrEmpty(ps.CurrentStyleSheet))
                        psv.SetCurrentStyleSheet(ps, ps.CurrentStyleSheet);

                    pi.OverrideSettings = ps;

                    // ✨ 修改：直接实例化 PlotInfoValidator 并验证
                    Autodesk.AutoCAD.PlottingServices.PlotInfoValidator piv = new Autodesk.AutoCAD.PlottingServices.PlotInfoValidator();
                    piv.MediaMatchingPolicy = Autodesk.AutoCAD.PlottingServices.MatchingPolicy.MatchEnabled;
                    piv.Validate(pi);

                    // 3. 执行打印过程
                    if (Autodesk.AutoCAD.PlottingServices.PlotFactory.ProcessPlotState == Autodesk.AutoCAD.PlottingServices.ProcessPlotState.NotPlotting)
                    {
                        using (var pe = Autodesk.AutoCAD.PlottingServices.PlotFactory.CreatePublishEngine())
                        {
                            pe.BeginPlot(null, null);
                            // ✨ 修改：修正 BeginDocument 签名，简化进度对话框参数
                            pe.BeginDocument(pi, doc.Name, null, 1, true, pdfPath);

                            Autodesk.AutoCAD.PlottingServices.PlotPageInfo ppi = new Autodesk.AutoCAD.PlottingServices.PlotPageInfo();
                            pe.BeginPage(ppi, pi, true, null);
                            pe.BeginGenerateGraphics(null);
                            pe.EndGenerateGraphics(null);
                            pe.EndPage(null);
                            pe.EndDocument(null);
                            pe.EndPlot(null);
                        }
                    }

                    // 4. 恢复图层状态
                    foreach (var ltr in layersToRestore) ltr.IsOff = false;

                    tr.Commit();
                }
            }
        }
        /// <summary>
        /// 增强版 SVG 导出：计算范围、导出数据，并在 CAD 中生成非打印的 DCFW 边界框。
        /// 修改：增加了创建 DCFW 图层及绘制矩形框的逻辑。
        /// </summary>
        public static Autodesk.AutoCAD.DatabaseServices.Extents3d ExportToSvg(IEnumerable<StyleMaster.Models.MaterialItem> items, string savePath)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return new Autodesk.AutoCAD.DatabaseServices.Extents3d();
            var db = doc.Database;
            var ed = doc.Editor;

            Autodesk.AutoCAD.DatabaseServices.Extents3d totalExt = new Autodesk.AutoCAD.DatabaseServices.Extents3d();

            try
            {
                using (doc.LockDocument())
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var btr = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)tr.GetObject(db.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        // 1. 预遍历计算范围
                        bool hasEnts = false;
                        foreach (var item in items)
                        {
                            var ids = GetEntitiesOnLayer(btr, tr, item.LayerName);
                            foreach (Autodesk.AutoCAD.DatabaseServices.ObjectId id in ids)
                            {
                                if (tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) is Autodesk.AutoCAD.DatabaseServices.Entity ent)
                                {
                                    if (!hasEnts) { totalExt = ent.GeometricExtents; hasEnts = true; }
                                    else { totalExt.AddExtents(ent.GeometricExtents); }
                                }
                            }
                        }

                        if (!hasEnts) return totalExt;

                        // --- ✨ 新增：创建 DCFW 图层并生成边界矩形 ---
                        string layerName = "DCFW";
                        Autodesk.AutoCAD.DatabaseServices.LayerTable lt = (Autodesk.AutoCAD.DatabaseServices.LayerTable)tr.GetObject(db.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        Autodesk.AutoCAD.DatabaseServices.ObjectId ltId;
                        if (!lt.Has(layerName))
                        {
                            Autodesk.AutoCAD.DatabaseServices.LayerTableRecord ltr = new Autodesk.AutoCAD.DatabaseServices.LayerTableRecord();
                            ltr.Name = layerName;
                            ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 4); // 青色
                            ltr.IsPlottable = false; // ✨ 设置为不可打印
                            ltId = lt.Add(ltr);
                            tr.AddNewlyCreatedDBObject(ltr, true);
                        }
                        else
                        {
                            ltId = lt[layerName];
                        }

                        // 在范围内生成矩形框
                        using (Autodesk.AutoCAD.DatabaseServices.Polyline rect = new Autodesk.AutoCAD.DatabaseServices.Polyline(4))
                        {
                            rect.AddVertexAt(0, new Autodesk.AutoCAD.Geometry.Point2d(totalExt.MinPoint.X, totalExt.MinPoint.Y), 0, 0, 0);
                            rect.AddVertexAt(1, new Autodesk.AutoCAD.Geometry.Point2d(totalExt.MaxPoint.X, totalExt.MinPoint.Y), 0, 0, 0);
                            rect.AddVertexAt(2, new Autodesk.AutoCAD.Geometry.Point2d(totalExt.MaxPoint.X, totalExt.MaxPoint.Y), 0, 0, 0);
                            rect.AddVertexAt(3, new Autodesk.AutoCAD.Geometry.Point2d(totalExt.MinPoint.X, totalExt.MaxPoint.Y), 0, 0, 0);
                            rect.Closed = true;
                            rect.LayerId = ltId;

                            btr.AppendEntity(rect);
                            tr.AddNewlyCreatedDBObject(rect, true);
                        }
                        // --- 边界框创建结束 ---

                        // 2. 导出 SVG 数据 (逻辑保持不变)
                        double width = totalExt.MaxPoint.X - totalExt.MinPoint.X;
                        double height = totalExt.MaxPoint.Y - totalExt.MinPoint.Y;
                        System.Text.StringBuilder svgContent = new System.Text.StringBuilder();
                        svgContent.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                        svgContent.AppendLine($"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"{totalExt.MinPoint.X} {-totalExt.MaxPoint.Y} {width} {height}\" data-width=\"{width:F2}\" data-height=\"{height:F2}\" data-minx=\"{totalExt.MinPoint.X:F2}\" data-miny=\"{totalExt.MinPoint.Y:F2}\">");

                        foreach (var item in items)
                        {
                            if (!item.IsFillLayer) continue; // 仅导出勾选了填充的层

                            var ids = GetEntitiesOnLayer(btr, tr, item.LayerName);
                            if (ids.Count == 0) continue;
                            svgContent.AppendLine($"  <g id=\"{item.LayerName}\" data-pattern=\"{item.PatternName}\">");
                            foreach (Autodesk.AutoCAD.DatabaseServices.ObjectId id in ids)
                            {
                                if (tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) is Autodesk.AutoCAD.DatabaseServices.Polyline pl)
                                {
                                    svgContent.Append("    <path d=\"M ");
                                    for (int i = 0; i < pl.NumberOfVertices; i++)
                                    {
                                        var pt = pl.GetPoint2dAt(i);
                                        svgContent.Append($"{pt.X} {-pt.Y} ");
                                        if (i < pl.NumberOfVertices - 1) svgContent.Append("L ");
                                    }
                                    svgContent.AppendLine("Z\" />");
                                }
                            }
                            svgContent.AppendLine("  </g>");
                        }
                        svgContent.AppendLine("</svg>");
                        System.IO.File.WriteAllText(savePath, svgContent.ToString());

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[错误] 导出失败: {ex.Message}");
            }

            return totalExt;
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
        /// <summary>
        /// 按层级数字从大到小（降序）执行批量填充预览。
        /// 逻辑：先填充数字大的图层，后填充数字小的图层。
        /// </summary>
        public static void RunFill(System.Collections.Generic.IEnumerable<StyleMaster.Models.MaterialItem> items)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var btr = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)tr.GetObject(db.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    // ✨ 核心修改：使用 OrderByDescending 实现降序排列
                    // 顺序将变为：4, 3, 2, 1 ...
                    var sortedItems = System.Linq.Enumerable.OrderByDescending(items, x => x.Priority);

                    foreach (var item in sortedItems)
                    {
                        // 获取该图层上的边界 (调用您现有的 GetEntitiesOnLayer 方法)
                        var boundaryIds = GetEntitiesOnLayer(btr, tr, item.LayerName);
                        if (boundaryIds.Count == 0) continue;

                        foreach (Autodesk.AutoCAD.DatabaseServices.ObjectId bId in boundaryIds)
                        {
                            using (var hatch = new Autodesk.AutoCAD.DatabaseServices.Hatch())
                            {
                                // 设置图案与属性
                                hatch.SetHatchPattern(Autodesk.AutoCAD.DatabaseServices.HatchPatternType.PreDefined, item.PatternName);
                                hatch.Layer = item.LayerName;
                                if (item.CadColor != null) hatch.Color = item.CadColor;

                                // 插入数据库
                                Autodesk.AutoCAD.DatabaseServices.ObjectId hId = btr.AppendEntity(hatch);
                                tr.AddNewlyCreatedDBObject(hatch, true);

                                // 关联边界
                                var idsCol = new Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection { bId };
                                hatch.AppendLoop(Autodesk.AutoCAD.DatabaseServices.HatchLoopTypes.Default, idsCol);
                                hatch.EvaluateHatch(true);

                                // ✨ 显式确保当前生成的填充位于最上方
                                // 这样后生成的（数字小的图层）会压在先生成的（数字大的图层）上面
                                using (var dot = (Autodesk.AutoCAD.DatabaseServices.DrawOrderTable)tr.GetObject(btr.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite))
                                {
                                    var hCol = new Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection { hId };
                                    dot.MoveToTop(hCol);
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
                // 强制刷新视图
                ed.Regen();
            }
        }
    }
}