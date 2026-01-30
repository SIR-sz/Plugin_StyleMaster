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
        /* * 文件位置：StyleMaster/Services/CadRenderingService.cs
 * 方法：PlotRepresentationPdf
 * 功能：自动化打印表现层线稿为 PDF。
 * 修改说明：增加了对 MaterialItem.IsExportEnabled 的判断。只有勾选了“打印”的图层才会显示，未勾选的图层在打印时会被临时关闭。
 */
        public static void PlotRepresentationPdf(Autodesk.AutoCAD.DatabaseServices.Extents3d ext, string pdfPath, System.Collections.Generic.IEnumerable<StyleMaster.Models.MaterialItem> items)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // 1. 根据用户在 UI 上的勾选状态，控制图层开关
                    var lt = (Autodesk.AutoCAD.DatabaseServices.LayerTable)tr.GetObject(db.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    var layersToRestore = new System.Collections.Generic.Dictionary<Autodesk.AutoCAD.DatabaseServices.LayerTableRecord, bool>();

                    foreach (var item in items)
                    {
                        if (lt.Has(item.LayerName))
                        {
                            var ltr = (Autodesk.AutoCAD.DatabaseServices.LayerTableRecord)tr.GetObject(lt[item.LayerName], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            // 备份原始状态
                            layersToRestore[ltr] = ltr.IsOff;

                            // 如果用户没勾选“打印”，则关闭图层
                            if (!item.IsExportEnabled)
                            {
                                ltr.IsOff = true;
                            }
                            else
                            {
                                ltr.IsOff = false;
                            }
                        }
                    }

                    // 2. 配置打印参数
                    var lo = (Autodesk.AutoCAD.DatabaseServices.Layout)tr.GetObject(Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current.GetLayoutId(Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current.CurrentLayout), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    Autodesk.AutoCAD.PlottingServices.PlotInfo pi = new Autodesk.AutoCAD.PlottingServices.PlotInfo { Layout = lo.ObjectId };

                    Autodesk.AutoCAD.DatabaseServices.PlotSettings ps = new Autodesk.AutoCAD.DatabaseServices.PlotSettings(lo.ModelType);
                    ps.CopyFrom(lo);
                    Autodesk.AutoCAD.DatabaseServices.PlotSettingsValidator psv = Autodesk.AutoCAD.DatabaseServices.PlotSettingsValidator.Current;

                    Autodesk.AutoCAD.DatabaseServices.Extents2d windowExt = new Autodesk.AutoCAD.DatabaseServices.Extents2d(ext.MinPoint.X, ext.MinPoint.Y, ext.MaxPoint.X, ext.MaxPoint.Y);
                    psv.SetPlotWindowArea(ps, windowExt);
                    psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
                    psv.SetUseStandardScale(ps, true);
                    psv.SetStdScaleType(ps, Autodesk.AutoCAD.DatabaseServices.StdScaleType.ScaleToFit);
                    psv.SetPlotCentered(ps, true);

                    // 强制使用 MediaBox 对应的全边距纸张
                    psv.SetPlotConfigurationName(ps, "AutoCAD PDF (High Quality Print).pc3", "ISO_full_bleed_A3_(420.00_x_297.00_MM)");

                    if (!string.IsNullOrEmpty(ps.CurrentStyleSheet))
                        psv.SetCurrentStyleSheet(ps, ps.CurrentStyleSheet);

                    pi.OverrideSettings = ps;

                    Autodesk.AutoCAD.PlottingServices.PlotInfoValidator piv = new Autodesk.AutoCAD.PlottingServices.PlotInfoValidator();
                    piv.MediaMatchingPolicy = Autodesk.AutoCAD.PlottingServices.MatchingPolicy.MatchEnabled;
                    piv.Validate(pi);

                    // 3. 执行打印
                    if (Autodesk.AutoCAD.PlottingServices.PlotFactory.ProcessPlotState == Autodesk.AutoCAD.PlottingServices.ProcessPlotState.NotPlotting)
                    {
                        using (var pe = Autodesk.AutoCAD.PlottingServices.PlotFactory.CreatePublishEngine())
                        {
                            pe.BeginPlot(null, null);
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

                    // 4. 恢复图层原始状态
                    foreach (var kvp in layersToRestore)
                    {
                        kvp.Key.IsOff = kvp.Value;
                    }

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
        /* * 文件位置：StyleMaster/Services/CadRenderingService.cs
  * 方法：ExportToSvg
  * 功能：导出包含颜色映射元数据的配置 SVG，用于 PS 脚本识别图层颜色
  */
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

                // 1. 计算选中的打印图层包围盒
                bool hasEnts = false;
                var exportItems = items.Where(x => x.IsExportEnabled).ToList();

                foreach (var item in exportItems)
                {
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

                // 2. 匹配模数对齐逻辑 (A3 比例)
                double rawW = totalExt.MaxPoint.X - totalExt.MinPoint.X;
                double rawH = totalExt.MaxPoint.Y - totalExt.MinPoint.Y;
                double maxDim = Math.Max(rawW, rawH);

                double baseA3Long = 420.0;
                double baseA3Short = 297.0;
                double modifier = 1.0;

                if (maxDim <= 0.42) modifier = 0.001;
                else if (maxDim <= 4.2) modifier = 0.01;
                else if (maxDim <= 42.0) modifier = 0.1;

                bool isLandscape = rawW > rawH;
                double unitW = (isLandscape ? baseA3Long : baseA3Short) * modifier;
                double unitH = (isLandscape ? baseA3Short : baseA3Long) * modifier;

                double targetW = Math.Max(1, Math.Ceiling(rawW / unitW)) * unitW;
                double targetH = Math.Max(1, Math.Ceiling(rawH / unitH)) * unitH;

                double centerX = (totalExt.MinPoint.X + totalExt.MaxPoint.X) / 2.0;
                double centerY = (totalExt.MinPoint.Y + totalExt.MaxPoint.Y) / 2.0;

                double dcfwMinX = centerX - (targetW / 2.0);
                double dcfwMaxY = centerY + (targetH / 2.0);
                double dcfwMinY = centerY - (targetH / 2.0);

                // 3. 生成辅助 DCFW 矩形
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

                // 4. 导出配置清单 SVG
                StringBuilder svg = new StringBuilder();
                svg.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                svg.AppendLine($"<svg xmlns=\"http://www.w3.org/2000/svg\" " +
                               $"data-width=\"{targetW}\" data-height=\"{targetH}\" " +
                               $"data-minx=\"{dcfwMinX}\" data-maxy=\"{dcfwMaxY}\">");

                foreach (var item in exportItems)
                {
                    // 获取图层的 HEX 颜色
                    System.Drawing.Color drawingCol = item.CadColor.ColorValue;
                    string hexColor = $"#{drawingCol.R:X2}{drawingCol.G:X2}{drawingCol.B:X2}";

                    // 仅导出元数据节点，不再导出 path 坐标
                    svg.AppendLine($"  <metadata id=\"{item.LayerName}\" color=\"{hexColor}\" />");
                }
                svg.AppendLine("</svg>");

                File.WriteAllText(savePath, svg.ToString());
                tr.Commit();

                ed.WriteMessage($"\n[StyleMaster] 配置导出成功! 请捕捉 DCFW 矩形打印 PDF。");

                // 返回修正后的 DCFW 范围
                return new Extents3d(new Point3d(dcfwMinX, dcfwMinY, 0), new Point3d(dcfwMinX + targetW, dcfwMaxY, 0));
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

        /* * 文件位置：StyleMaster/Services/CadRenderingService.cs
 * 方法：GetEntitiesOnLayer
 * 功能：获取指定图层上的边界实体。
 * 修改说明：增加了对 BlockReference 的过滤，防止外部参照干扰填充计算，从而修复 eInvalidInput 报错。
 */
        private static ObjectIdCollection GetEntitiesOnLayer(BlockTableRecord btr, Transaction tr, string layerName)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            foreach (ObjectId id in btr)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent != null && ent.Layer.Equals(layerName, StringComparison.OrdinalIgnoreCase))
                {
                    // 核心过滤逻辑：排除填充、图片以及块参照（外部参照）
                    if (!(ent is Hatch) &&
                        !(ent is RasterImage) &&
                        !(ent is BlockReference)) // 显式过滤外部参照和普通块
                    {
                        ids.Add(id);
                    }
                }
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