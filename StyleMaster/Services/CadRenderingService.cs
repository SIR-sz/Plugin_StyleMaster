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
        /// 先删除该图层已有的填充，再根据当前设置项重新生成。
        /// </summary>
        /// <param name="item">需要刷新的材质设置项</param>
        public static void RefreshSingleLayer(StyleMaster.Models.MaterialItem item)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (doc.LockDocument())
            {
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var btr = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)tr.GetObject(doc.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    // 1. 删除该图层上现有的所有填充
                    foreach (Autodesk.AutoCAD.DatabaseServices.ObjectId id in btr)
                    {
                        var ent = tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                        if (ent is Autodesk.AutoCAD.DatabaseServices.Hatch hatch && hatch.Layer.Equals(item.LayerName, StringComparison.OrdinalIgnoreCase))
                        {
                            tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                            ent.Erase();
                        }
                    }

                    // 2. 重新扫描边界并创建新填充
                    var ids = GetEntitiesOnLayer(btr, tr, item.LayerName);
                    foreach (Autodesk.AutoCAD.DatabaseServices.ObjectId boundaryId in ids)
                    {
                        if (item.FillMode == StyleMaster.Models.FillType.Hatch)
                        {
                            CreateHatch(tr, btr, boundaryId, item);
                        }
                    }

                    tr.Commit();
                    doc.Editor.WriteMessage($"\n[StyleMaster] 图层 {item.LayerName} 刷新完成。");
                }
            }
        }
        /// <summary>
        /// 执行全局一键填充逻辑。
        /// 按照优先级从底层到顶层遍历图层，并对指定图层内的多段线进行填充。
        /// </summary>
        /// <param name="items">材质映射项集合</param>
        public static void ExecuteFill(IEnumerable<MaterialItem> items)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var db = doc.Database;
            var ed = doc.Editor;

            // ✨ 核心修正：非模态窗口必须锁定文档才能安全写入数据库数据
            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // 获取当前空间（模型空间或布局空间）
                    var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    int totalFilled = 0;

                    // 按照 Priority 降序排列（数字大的先填，作为底层背景）
                    var sortedItems = items.OrderByDescending(x => x.Priority).ToList();

                    foreach (var item in sortedItems)
                    {
                        // 获取目标图层上的多段线实体
                        var ids = GetEntitiesOnLayer(btr, tr, item.LayerName);

                        if (ids.Count > 0)
                        {
                            ed.WriteMessage($"\n[StyleMaster] 正在处理图层: {item.LayerName}，发现 {ids.Count} 个对象...");

                            foreach (ObjectId id in ids)
                            {
                                if (item.FillMode == FillType.Hatch)
                                {
                                    CreateHatch(tr, btr, id, item);
                                    totalFilled++;
                                }
                            }
                        }
                    }

                    tr.Commit();
                    ed.WriteMessage($"\n[StyleMaster] 一键填充操作完成！共生成 {totalFilled} 个填充对象。");
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
    }
}