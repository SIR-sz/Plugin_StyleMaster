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
        /// <summary>
        /// 增强版 SVG 导出：匹配标准纸张，生成辅助矩形，并导出对齐元数据。
        /// </summary>
        public static Extents3d ExportToSvg(IEnumerable<MaterialItem> items, string savePath)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            Extents3d totalExt = new Extents3d();

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                // 1. 计算原始内容包围盒
                bool hasEnts = false;
                foreach (var item in items)
                {
                    if (!item.IsFillLayer) continue;
                    var ids = GetEntitiesOnLayer(btr, tr, item.LayerName);
                    foreach (ObjectId id in ids)
                    {
                        if (tr.GetObject(id, OpenMode.ForRead) is Entity ent)
                        {
                            if (!hasEnts) { totalExt = ent.GeometricExtents; hasEnts = true; }
                            else { totalExt.AddExtents(ent.GeometricExtents); }
                        }
                    }
                }

                if (!hasEnts) return totalExt;

                // 2. 动态匹配 A3 比例模数 (向下兼容 0.1, 0.01, 0.001 倍)
                double rawW = totalExt.MaxPoint.X - totalExt.MinPoint.X;
                double rawH = totalExt.MaxPoint.Y - totalExt.MinPoint.Y;
                double maxDim = Math.Max(rawW, rawH);

                // 定义基础 A3 尺寸
                double baseA3Long = 420.0;
                double baseA3Short = 297.0;

                // 自动寻找模数阶梯
                double modifier = 1.0;
                if (maxDim <= 0.42) modifier = 0.001;      // 针对毫米级微型详图
                else if (maxDim <= 4.2) modifier = 0.01;   // 针对厘米级
                else if (maxDim <= 42.0) modifier = 0.1;   // 针对分米级/较小构件
                else modifier = 1.0;                       // 标准 A3 或以上

                bool isLandscape = rawW > rawH;
                double unitW = (isLandscape ? baseA3Long : baseA3Short) * modifier;
                double unitH = (isLandscape ? baseA3Short : baseA3Long) * modifier;

                // 计算整数倍数 (向上取整)
                double multiW = Math.Max(1, Math.Ceiling(rawW / unitW));
                double multiH = Math.Max(1, Math.Ceiling(rawH / unitH));

                double targetW = multiW * unitW;
                double targetH = multiH * unitH;

                // 3. 计算中心对齐的辅助边界 (DCFW)
                double centerX = (totalExt.MinPoint.X + totalExt.MaxPoint.X) / 2.0;
                double centerY = (totalExt.MinPoint.Y + totalExt.MaxPoint.Y) / 2.0;

                double dcfwMinX = centerX - (targetW / 2.0);
                double dcfwMaxY = centerY + (targetH / 2.0); // 顶部
                double dcfwMinY = centerY - (targetH / 2.0); // 底部

                // 4. 更新 DCFW 矩形框
                string dcfwLayer = "DCFW";
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
                if (!lt.Has(dcfwLayer))
                {
                    var ltr = new LayerTableRecord { Name = dcfwLayer, Color = Color.FromColorIndex(ColorMethod.ByAci, 4), IsPlottable = false };
                    lt.Add(ltr); tr.AddNewlyCreatedDBObject(ltr, true);
                }

                using (Polyline rect = new Polyline(4))
                {
                    rect.AddVertexAt(0, new Point2d(dcfwMinX, dcfwMinY), 0, 0, 0);
                    rect.AddVertexAt(1, new Point2d(dcfwMinX + targetW, dcfwMinY), 0, 0, 0);
                    rect.AddVertexAt(2, new Point2d(dcfwMinX + targetW, dcfwMaxY), 0, 0, 0);
                    rect.AddVertexAt(3, new Point2d(dcfwMinX, dcfwMaxY), 0, 0, 0);
                    rect.Closed = true; rect.Layer = dcfwLayer;
                    btr.AppendEntity(rect); tr.AddNewlyCreatedDBObject(rect, true);
                }

                // 5. 导出 SVG 并写入优化后的坐标元数据
                // 5. 导出 SVG 并写入优化后的坐标元数据
                StringBuilder svg = new StringBuilder();
                svg.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                // 使用标准的 viewBox 格式，并存储 data 属性供 PS 使用
                svg.AppendLine($"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"{dcfwMinX} {-dcfwMaxY} {targetW} {targetH}\" " +
                               $"data-width=\"{targetW}\" data-height=\"{targetH}\" " +
                               $"data-minx=\"{dcfwMinX}\" data-maxy=\"{dcfwMaxY}\">");

                foreach (var item in items)
                {
                    if (!item.IsFillLayer) continue;
                    var ids = GetEntitiesOnLayer(btr, tr, item.LayerName);
                    svg.AppendLine($"  <g id=\"{item.LayerName}\">");
                    foreach (ObjectId id in ids)
                    {
                        if (tr.GetObject(id, OpenMode.ForRead) is Polyline pl)
                        {
                            svg.Append("    <path d=\"M ");
                            for (int i = 0; i < pl.NumberOfVertices; i++)
                            {
                                Point2d pt = pl.GetPoint2dAt(i);
                                // 统一 Y 轴取反逻辑
                                svg.Append($"{pt.X:F4} {-pt.Y:F4} ");
                                if (i < pl.NumberOfVertices - 1) svg.Append("L ");
                            }
                            // 修改点 1：将 fill="none" 改为 fill="black"，方便 PS 脚本识别闭合区域生成选区
                            svg.AppendLine("Z\" fill=\"black\" stroke=\"none\" />");
                        }
                    }
                    svg.AppendLine("  </g>");
                }
                svg.AppendLine("</svg>");

                File.WriteAllText(savePath, svg.ToString());
                tr.Commit();

                // 修改点 2：构造基于 DCFW 辅助矩形的 Extents3d 用于返回
                Extents3d dcfwExt = new Extents3d(
                    new Point3d(dcfwMinX, dcfwMinY, 0),
                    new Point3d(dcfwMinX + targetW, dcfwMaxY, 0)
                );

                ed.WriteMessage($"\n[StyleMaster] 导出成功! 匹配模数: {modifier}x A3.");
                ed.WriteMessage($"\n[提示] 请捕捉 DCFW 矩形窗口打印 PDF，确保与 SVG 完美重叠。");

                // 返回修正后的范围，确保 UI 层输出的坐标与 PDF 窗口一致
                return dcfwExt;
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
        /// 创建单个填充对象
        /// 修改点：移除了 settings.Opacity 引用，增加了 DrawOrder 置顶逻辑
        /// </summary>
        private static void CreateHatch(Autodesk.AutoCAD.DatabaseServices.Transaction tr, Autodesk.AutoCAD.DatabaseServices.BlockTableRecord btr, Autodesk.AutoCAD.DatabaseServices.ObjectId boundaryId, StyleMaster.Models.MaterialItem settings, Autodesk.AutoCAD.EditorInput.Editor ed)
        {
            try
            {
                string patternName = string.IsNullOrEmpty(settings.PatternName) ? "ANSI31" : settings.PatternName;

                Autodesk.AutoCAD.DatabaseServices.Hatch hat = new Autodesk.AutoCAD.DatabaseServices.Hatch();
                hat.SetDatabaseDefaults();
                hat.Layer = settings.LayerName;

                // 将填充添加到块表记录
                Autodesk.AutoCAD.DatabaseServices.ObjectId hatId = btr.AppendEntity(hat);
                tr.AddNewlyCreatedDBObject(hat, true);

                hat.Associative = true;
                hat.Normal = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis;

                if (settings.CadColor != null)
                {
                    hat.Color = settings.CadColor;
                }

                // 设置图案与比例
                try
                {
                    hat.SetHatchPattern(Autodesk.AutoCAD.DatabaseServices.HatchPatternType.PreDefined, patternName);
                    hat.PatternScale = settings.Scale > 0 ? settings.Scale : 1.0;
                    // 重新设置以触发比例刷新
                    hat.SetHatchPattern(hat.PatternType, hat.PatternName);
                }
                catch
                {
                    // 如果图案名无效，降级使用 SOLID 填充
                    hat.SetHatchPattern(Autodesk.AutoCAD.DatabaseServices.HatchPatternType.PreDefined, "SOLID");
                }

                // ✨ 修改：删除了之前导致报错的 hat.Transparency 设置行

                // 添加边界循环并计算
                Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection idsCol = new Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection { boundaryId };
                hat.AppendLoop(Autodesk.AutoCAD.DatabaseServices.HatchLoopTypes.Outermost, idsCol);
                hat.EvaluateHatch(true);

                // ✨ 新增：强制设置绘图次序，确保新生成的对象在视觉最上方
                using (var dot = (Autodesk.AutoCAD.DatabaseServices.DrawOrderTable)tr.GetObject(btr.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite))
                {
                    Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection hIds = new Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection { hatId };
                    dot.MoveToTop(hIds);
                }
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
        /// <summary>
        /// 按层级数字从大到小（降序）执行批量填充预览。
        /// 修改：移除了对 Opacity 属性的引用。
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

                    // 按 Priority 降序排列 (4, 3, 2, 1)
                    var sortedItems = System.Linq.Enumerable.OrderByDescending(items, x => x.Priority);

                    foreach (var item in sortedItems)
                    {
                        var boundaryIds = GetEntitiesOnLayer(btr, tr, item.LayerName);
                        if (boundaryIds.Count == 0) continue;

                        foreach (Autodesk.AutoCAD.DatabaseServices.ObjectId bId in boundaryIds)
                        {
                            using (var hatch = new Autodesk.AutoCAD.DatabaseServices.Hatch())
                            {
                                hatch.SetHatchPattern(Autodesk.AutoCAD.DatabaseServices.HatchPatternType.PreDefined, item.PatternName);
                                hatch.Layer = item.LayerName;
                                if (item.CadColor != null) hatch.Color = item.CadColor;

                                // ✨ 修改：删除了设置 hatch.Transparency 的逻辑

                                btr.AppendEntity(hatch);
                                tr.AddNewlyCreatedDBObject(hatch, true);

                                var idsCol = new Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection { bId };
                                hatch.AppendLoop(Autodesk.AutoCAD.DatabaseServices.HatchLoopTypes.Default, idsCol);
                                hatch.EvaluateHatch(true);

                                using (var dot = (Autodesk.AutoCAD.DatabaseServices.DrawOrderTable)tr.GetObject(btr.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite))
                                {
                                    var hCol = new Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection { hatch.ObjectId };
                                    dot.MoveToTop(hCol);
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
                ed.Regen();
            }
        }
    }
}