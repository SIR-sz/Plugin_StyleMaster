using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using StyleMaster.Models;
using System;
using System.Collections.Generic;
using System.Linq;



namespace StyleMaster.Services
{
    public class CadRenderingService
    {
        /// <summary>
        /// 清除指定图层列表上的所有 Hatch 填充实体。
        /// </summary>
        /// <param name="layerNames">需要清理的图层名称集合</param>
        public static void ClearFillsOnLayers(System.Collections.Generic.IEnumerable<string> layerNames)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (doc.LockDocument())
            {
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var btr = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)tr.GetObject(doc.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                    int count = 0;

                    foreach (Autodesk.AutoCAD.DatabaseServices.ObjectId id in btr)
                    {
                        var ent = tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                        if (ent is Autodesk.AutoCAD.DatabaseServices.Hatch hatch && layerNames.Any(l => l.Equals(hatch.Layer, StringComparison.OrdinalIgnoreCase)))
                        {
                            tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                            ent.Erase();
                            count++;
                        }
                    }
                    tr.Commit();
                    doc.Editor.WriteMessage($"\n[StyleMaster] 已成功清除 {count} 个填充对象。");
                }
            }
        }

        /// <summary>
        /// 刷新单个图层的填充逻辑。
        /// 先删除该图层已有的填充，重新生成后，根据全局列表顺序修正绘图次序。
        /// </summary>
        /// <param name="item">需要刷新的材质设置项</param>
        /// <param name="allItems">用于层级参考的完整材质列表</param>
        public static void RefreshSingleLayer(MaterialItem item, IEnumerable<MaterialItem> allItems)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (doc.LockDocument())
            {
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var btr = (BlockTableRecord)tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);

                    // 1. 删除该图层上现有的所有填充
                    foreach (ObjectId id in btr)
                    {
                        var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent is Hatch hatch && hatch.Layer.Equals(item.LayerName, StringComparison.OrdinalIgnoreCase))
                        {
                            tr.GetObject(id, OpenMode.ForWrite);
                            ent.Erase();
                        }
                    }

                    // 2. 重新扫描边界并创建新填充
                    var ids = GetEntitiesOnLayer(btr, tr, item.LayerName);
                    foreach (ObjectId boundaryId in ids)
                    {
                        if (item.FillMode == FillType.Hatch)
                        {
                            CreateHatch(tr, btr, boundaryId, item);
                        }
                    }

                    // 3. ✨ 关键修正：重新根据全局列表排序，维持正确的层级感
                    ReorderDrawOrder(tr, btr, allItems);

                    tr.Commit();
                    doc.Editor.WriteMessage($"\n[StyleMaster] 图层 {item.LayerName} 已刷新并修正层级。");
                }
            }
        }
        /// <summary>
        /// 执行全量同步填充（幂等性操作）。
        /// 逻辑：锁定文档 -> 清除列表内图层已有的填充 -> 重新生成填充 -> 锁定绘图次序。
        /// </summary>
        /// <param name="items">当前的材质配置列表</param>
        public static void ExecuteFill(System.Collections.Generic.IEnumerable<StyleMaster.Models.MaterialItem> items)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var db = doc.Database;
            var ed = doc.Editor;

            // 获取当前所有参与配置的图层名单
            var targetLayerNames = items.Select(x => x.LayerName).ToList();

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var btr = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)tr.GetObject(db.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    // --- 第一步：清理阶段 ---
                    int deletedCount = 0;
                    foreach (Autodesk.AutoCAD.DatabaseServices.ObjectId id in btr)
                    {
                        var ent = tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                        if (ent is Autodesk.AutoCAD.DatabaseServices.Hatch hatch)
                        {
                            // 如果该填充所在的图层在我们的配置列表中，则删除它
                            if (targetLayerNames.Any(name => name.Equals(hatch.Layer, System.StringComparison.OrdinalIgnoreCase)))
                            {
                                tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                ent.Erase();
                                deletedCount++;
                            }
                        }
                    }

                    // --- 第二步：填充阶段 ---
                    int totalFilled = 0;
                    foreach (var item in items)
                    {
                        // 获取该图层上的边界对象（多段线等）
                        var boundaryIds = GetEntitiesOnLayer(btr, tr, item.LayerName);
                        if (boundaryIds.Count > 0)
                        {
                            foreach (Autodesk.AutoCAD.DatabaseServices.ObjectId boundaryId in boundaryIds)
                            {
                                if (item.FillMode == StyleMaster.Models.FillType.Hatch)
                                {
                                    CreateHatch(tr, btr, boundaryId, item);
                                    totalFilled++;
                                }
                            }
                        }
                    }

                    // --- 第三步：层级修正阶段 ---
                    // 调用之前写的 ReorderDrawOrder 方法，确保叠加顺序正确
                    ReorderDrawOrder(tr, btr, items);

                    tr.Commit();
                    ed.WriteMessage($"\n[StyleMaster] 同步完成：清理了 {deletedCount} 个旧填充，新生成了 {totalFilled} 个填充并锁定层级。");
                }
            }
        }

        private static List<ObjectId> GetEntitiesOnLayer(BlockTableRecord btr, Transaction tr, string layer)
        {
            var list = new List<ObjectId>();
            foreach (ObjectId id in btr)
            {
                var ent = (Entity)tr.GetObject(id, OpenMode.ForRead);
                // 匹配图层名（忽略大小写）且必须是曲线类对象（多段线属于 Curve）
                if (ent.Layer.Equals(layer, StringComparison.OrdinalIgnoreCase) && ent is Curve)
                {
                    list.Add(id);
                }
            }
            return list;
        }
        /// <summary>
        /// 在指定的边界内创建 Hatch 填充对象。
        /// 应用用户选定的 AutoCAD 原生颜色，并取消背景色逻辑。
        /// </summary>
        private static void CreateHatch(Transaction tr, BlockTableRecord btr, ObjectId boundaryId, MaterialItem settings)
        {
            Hatch hat = new Hatch();
            hat.SetDatabaseDefaults();

            // 设置图层
            hat.Layer = settings.LayerName;

            btr.AppendEntity(hat);
            tr.AddNewlyCreatedDBObject(hat, true);

            // 设置标高与法线
            hat.Elevation = 0.0;
            hat.Normal = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis;

            // ✨ 核心修正：应用用户选择的颜色
            hat.Color = settings.CadColor;

            // 设置图案与比例
            hat.SetHatchPattern(HatchPatternType.PreDefined, settings.PatternName);
            hat.PatternScale = settings.Scale;

            // 设置透明度
            hat.Transparency = new Transparency((byte)(255 * (1 - settings.Opacity / 100.0)));

            // 设置边界
            ObjectIdCollection ids = new ObjectIdCollection { boundaryId };
            hat.Associative = true;
            hat.AppendLoop(HatchLoopTypes.Outermost, ids);

            // 评估填充
            hat.EvaluateHatch(true);

            // 建议方向：在这里可以添加 DrawOrderTable 逻辑将填充置于底层
        }
        /// <summary>
        /// 根据材质列表的顺序，重新排列所有填充对象的绘图次序（Draw Order）。
        /// 确保列表顶部的图层在视觉上也处于最前端。
        /// </summary>
        /// <param name="tr">当前事务</param>
        /// <param name="btr">当前空间块表记录</param>
        /// <param name="items">当前的材质配置列表</param>
        private static void ReorderDrawOrder(Transaction tr, BlockTableRecord btr, IEnumerable<MaterialItem> items)
        {
            // 获取当前空间的绘图次序表
            var dotId = btr.DrawOrderTableId;
            if (dotId.IsNull) return;

            var dot = (DrawOrderTable)tr.GetObject(dotId, OpenMode.ForWrite);

            // 按照 Priority 从大到小排列（列表底部的先处理，放置在底层；列表顶部的后处理，推至顶层）
            var sortedItems = items.OrderByDescending(x => x.Priority).ToList();

            foreach (var item in sortedItems)
            {
                ObjectIdCollection idsToMove = new ObjectIdCollection();

                // 搜寻当前空间中属于该图层的所有填充对象
                foreach (ObjectId id in btr)
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent is Hatch hatch && hatch.Layer.Equals(item.LayerName, StringComparison.OrdinalIgnoreCase))
                    {
                        idsToMove.Add(id);
                    }
                }

                // 将该图层的所有填充移动到最上方
                if (idsToMove.Count > 0)
                {
                    dot.MoveToTop(idsToMove);
                }
            }
        }
    }
}