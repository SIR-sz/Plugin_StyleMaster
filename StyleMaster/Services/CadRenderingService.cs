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
        // 在 StyleMaster/Services/CadRenderingService.cs 中
        private static void CreateHatch(Autodesk.AutoCAD.DatabaseServices.Transaction tr, BlockTableRecord btr, ObjectId boundaryId, MaterialItem settings)
        {
            Hatch hat = new Hatch();
            hat.SetDatabaseDefaults();

            // 1. 设置图层：确保填充生成在对应的图层上
            hat.Layer = settings.LayerName;

            btr.AppendEntity(hat);

            // 2. 修复：使用 AddNewlyCreatedDBObject (注意中间的 DB)
            tr.AddNewlyCreatedDBObject(hat, true);

            // 3. 修复：将 ZCoordinate 修改为 Elevation
            hat.Elevation = 0.0;
            hat.Normal = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis; // 确保法线方向正确

            // 4. 设置图案和比例 (注意 PreDefined 的大写 D)
            hat.SetHatchPattern(HatchPatternType.PreDefined, settings.PatternName);
            hat.PatternScale = settings.Scale;

            // 5. 设置透明度
            hat.Transparency = new Autodesk.AutoCAD.Colors.Transparency((byte)(255 * (1 - settings.Opacity / 100.0)));

            // 6. 设置边界并关联
            ObjectIdCollection ids = new ObjectIdCollection { boundaryId };
            hat.Associative = true;
            hat.AppendLoop(HatchLoopTypes.Outermost, ids); // 使用 Outermost 边界类型更稳定

            // 7. 关键：计算填充
            hat.EvaluateHatch(true);
        }
    }
}