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
        public static void ExecuteFill(IEnumerable<MaterialItem> items)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var sortedItems = items.OrderByDescending(x => x.Priority).ToList();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                foreach (var item in sortedItems)
                {
                    // 这里需要根据图层获取实体的逻辑
                    var ids = GetEntitiesOnLayer(db, tr, item.LayerName);
                    if (ids.Count == 0) continue;

                    foreach (ObjectId id in ids)
                    {
                        if (item.FillMode == FillType.Hatch)
                        {
                            CreateHatch(tr, btr, id, item);
                        }
                    }
                }
                tr.Commit();
            }
        }

        private static List<ObjectId> GetEntitiesOnLayer(Database db, Transaction tr, string layerName)
        {
            var result = new List<ObjectId>();
            BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
            foreach (ObjectId id in btr)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent != null && ent.Layer.Equals(layerName, StringComparison.OrdinalIgnoreCase))
                {
                    if (ent is Polyline || ent is Polyline2d) result.Add(id);
                }
            }
            return result;
        }

        private static void CreateHatch(Transaction tr, BlockTableRecord btr, ObjectId boundaryId, MaterialItem settings)
        {
            Hatch hat = new Hatch();
            hat.SetDatabaseDefaults(); // 建议加上，初始化默认值
            btr.AppendEntity(hat);
            tr.AddNewlyCreatedObject(hat, true);

            // ✨ 修复：使用 PreDefined 而非 Predefined
            hat.SetHatchPattern(HatchPatternType.PreDefined, settings.PatternName);
            hat.PatternScale = settings.Scale;

            // 处理透明度
            hat.Transparency = new Transparency((byte)(255 * (1 - settings.Opacity / 100.0)));

            ObjectIdCollection ids = new ObjectIdCollection();
            ids.Add(boundaryId);
            hat.AppendLoop(HatchLoopTypes.Default, ids);
            hat.EvaluateHatch(true);
        }
    }
}