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
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                foreach (var item in items.OrderByDescending(x => x.Priority))
                {
                    var ids = GetEntitiesOnLayer(btr, tr, item.LayerName);
                    foreach (ObjectId id in ids)
                    {
                        if (item.FillMode == FillType.Hatch) CreateHatch(tr, btr, id, item);
                    }
                }
                tr.Commit();
            }
        }

        private static List<ObjectId> GetEntitiesOnLayer(BlockTableRecord btr, Transaction tr, string layer)
        {
            var list = new List<ObjectId>();
            foreach (ObjectId id in btr)
            {
                var ent = (Entity)tr.GetObject(id, OpenMode.ForRead);
                if (ent.Layer.Equals(layer, System.StringComparison.OrdinalIgnoreCase) && ent is Curve)
                    list.Add(id);
            }
            return list;
        }

        // 在 StyleMaster/Services/CadRenderingService.cs 中
        private static void CreateHatch(Autodesk.AutoCAD.DatabaseServices.Transaction tr, BlockTableRecord btr, ObjectId boundaryId, MaterialItem settings)
        {
            Hatch hat = new Hatch();
            hat.SetDatabaseDefaults();
            btr.AppendEntity(hat);

            // 确保使用全限定名以解决 CS1061 错误
            tr.AddNewlyCreatedObject(hat, true);

            // 修复之前的 PreDefined 拼写问题（注意 D 为大写）
            hat.SetHatchPattern(HatchPatternType.PreDefined, settings.PatternName);
            hat.PatternScale = settings.Scale;
            hat.Transparency = new Transparency((byte)(255 * (1 - settings.Opacity / 100.0)));

            ObjectIdCollection ids = new ObjectIdCollection();
            ids.Add(boundaryId);
            hat.AppendLoop(HatchLoopTypes.Default, ids);
            hat.EvaluateHatch(true);
        }
    }
}