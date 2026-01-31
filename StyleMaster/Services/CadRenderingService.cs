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
        /// 强化版：通过强制设置 WCS 坐标系和法线向量，解决大坐标或 UCS 导致的 eInvalidInput 错误。
        /// </summary>
        private static void CreateHatch(Autodesk.AutoCAD.DatabaseServices.Transaction tr, Autodesk.AutoCAD.DatabaseServices.BlockTableRecord btr, Autodesk.AutoCAD.DatabaseServices.ObjectId boundaryId, StyleMaster.Models.MaterialItem settings, Autodesk.AutoCAD.EditorInput.Editor ed)
        {
            try
            {
                // 1. 获取边界实体并进行基础校验
                var ent = tr.GetObject(boundaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                if (ent == null) return;

                // 2. 初始化 Hatch 对象
                string patternName = string.IsNullOrEmpty(settings.PatternName) ? "SOLID" : settings.PatternName;
                using (Autodesk.AutoCAD.DatabaseServices.Hatch hat = new Autodesk.AutoCAD.DatabaseServices.Hatch())
                {
                    hat.SetDatabaseDefaults();
                    hat.Layer = settings.LayerName;

                    // 核心修复：强制 Hatch 使用世界坐标系的法线
                    hat.Normal = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis;
                    hat.Elevation = 0.0; // 强制填充平面在 Z=0

                    // 将填充添加到块表记录
                    btr.AppendEntity(hat);
                    tr.AddNewlyCreatedDBObject(hat, true);

                    // 设置图案
                    try
                    {
                        hat.SetHatchPattern(Autodesk.AutoCAD.DatabaseServices.HatchPatternType.PreDefined, patternName);
                        hat.PatternScale = settings.Scale > 0 ? settings.Scale : 1.0;
                    }
                    catch
                    {
                        hat.SetHatchPattern(Autodesk.AutoCAD.DatabaseServices.HatchPatternType.PreDefined, "SOLID");
                    }

                    if (settings.CadColor != null)
                    {
                        hat.Color = settings.CadColor;
                    }

                    // 核心修复：构造边界循环
                    Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection idsCol = new Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection { boundaryId };

                    // 使用 Try-Catch 包裹边界添加过程，防止单个边界错误中断整个图层填充
                    try
                    {
                        hat.Associative = true;
                        hat.AppendLoop(Autodesk.AutoCAD.DatabaseServices.HatchLoopTypes.Default, idsCol);
                        hat.EvaluateHatch(true);
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        ed.WriteMessage($"\n  [跳过] 图层 {settings.LayerName} 上的一个边界无效: {ex.Message}");
                        hat.Erase(); // 如果填充失败，删除无效的 Hatch 对象
                        return;
                    }

                    // 设置绘图次序
                    using (var dot = (Autodesk.AutoCAD.DatabaseServices.DrawOrderTable)tr.GetObject(btr.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite))
                    {
                        Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection hIds = new Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection { hat.ObjectId };
                        dot.MoveToTop(hIds);
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n  [错误] 填充逻辑崩溃: {ex.Message}");
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
        /// 强化版：增加了对坐标系和实体有效性的前置检查。
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
                            // 调用强化后的私有填充方法
                            CreateHatch(tr, btr, bId, item, ed);
                        }
                    }
                    tr.Commit();
                }
                ed.Regen();
            }
        }
        /// <summary>
        /// 将材质配置列表序列化为 XML 字符串并保存到当前 CAD 图纸的命名对象字典 (NOD) 中。
        /// 使用 AutoCAD 原生支持的存储方式，无需外部 DLL 依赖。
        /// </summary>
        public static void SaveToDatabase(System.Collections.ObjectModel.ObservableCollection<StyleMaster.Models.MaterialItem> items)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var nod = (Autodesk.AutoCAD.DatabaseServices.DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    string dictName = "StyleMaster_Data_Xml";
                    Autodesk.AutoCAD.DatabaseServices.Xrecord xrec;

                    if (nod.Contains(dictName))
                    {
                        xrec = (Autodesk.AutoCAD.DatabaseServices.Xrecord)tr.GetObject(nod.GetAt(dictName), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                    }
                    else
                    {
                        xrec = new Autodesk.AutoCAD.DatabaseServices.Xrecord();
                        nod.SetAt(dictName, xrec);
                        tr.AddNewlyCreatedDBObject(xrec, true);
                    }

                    // 使用 XmlSerializer 进行序列化
                    string xmlString = string.Empty;
                    try
                    {
                        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(System.Collections.Generic.List<StyleMaster.Models.MaterialItem>));
                        using (var sw = new System.IO.StringWriter())
                        {
                            serializer.Serialize(sw, System.Linq.Enumerable.ToList(items));
                            xmlString = sw.ToString();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        doc.Editor.WriteMessage($"\n[StyleMaster] 序列化失败: {ex.Message}");
                        return;
                    }

                    // 将 XML 字符串存入 Xrecord
                    // 使用 DxfCode.Text (1) 存储字符串
                    xrec.Data = new Autodesk.AutoCAD.DatabaseServices.ResultBuffer(
                        new Autodesk.AutoCAD.DatabaseServices.TypedValue((int)Autodesk.AutoCAD.DatabaseServices.DxfCode.Text, xmlString)
                    );

                    tr.Commit();
                }
            }
        }
        /* * 文件位置：StyleMaster/Services/CadRenderingService.cs
 * 方法：LoadFromDatabase
 * 功能：从 NOD 字典加载配置。仅还原数据属性，视觉还原交由 UI 层处理。
 */
        public static System.Collections.ObjectModel.ObservableCollection<StyleMaster.Models.MaterialItem> LoadFromDatabase()
        {
            var result = new System.Collections.ObjectModel.ObservableCollection<StyleMaster.Models.MaterialItem>();
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return result;

            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                var nod = (Autodesk.AutoCAD.DatabaseServices.DBDictionary)tr.GetObject(doc.Database.NamedObjectsDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                string dictName = "StyleMaster_Data_Xml";

                if (nod.Contains(dictName))
                {
                    var xrec = (Autodesk.AutoCAD.DatabaseServices.Xrecord)tr.GetObject(nod.GetAt(dictName), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    var rb = xrec.Data;
                    if (rb != null)
                    {
                        foreach (var tv in rb.AsArray())
                        {
                            if (tv.TypeCode == (int)Autodesk.AutoCAD.DatabaseServices.DxfCode.Text)
                            {
                                string xmlString = (string)tv.Value;
                                try
                                {
                                    var serializer = new System.Xml.Serialization.XmlSerializer(typeof(System.Collections.Generic.List<StyleMaster.Models.MaterialItem>));
                                    using (var sr = new System.IO.StringReader(xmlString))
                                    {
                                        var items = (System.Collections.Generic.List<StyleMaster.Models.MaterialItem>)serializer.Deserialize(sr);
                                        if (items != null)
                                        {
                                            foreach (var item in items) result.Add(item);
                                        }
                                    }
                                }
                                catch { }
                                break;
                            }
                        }
                    }
                }
                tr.Commit();
            }
            return result;
        }
    }
}