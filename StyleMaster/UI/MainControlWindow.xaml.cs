using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Microsoft.Win32; // <--- 修复 SaveFileDialog 错误的关键
using StyleMaster.Models;
using StyleMaster.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;


namespace StyleMaster.UI
{

    /// <summary>
    /// MainControlWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainControlWindow : Window
    {
        private Point _dragStartPoint;
        private MaterialItem _draggedItem;
        private bool _isDraggingNow = false;
        private static MainControlWindow _instance;

        // 数据源集合
        public ObservableCollection<MaterialItem> MaterialItems { get; set; }

        /// <summary>
        /// 构造函数：执行 UI 初始化，并设置数据绑定上下文。
        /// </summary>
        public MainControlWindow()
        {
            InitializeComponent();

            MaterialItems = new ObservableCollection<MaterialItem>();

            // ✨ 修改 1：在初始化集合后，立即注册监听器
            RegisterLayerPropertyTracker();

            InitializeTestData();

            this.DataContext = this;
            this.MainDataGrid.ItemsSource = MaterialItems;
        }
        /// <summary>
        /// 注册属性追踪逻辑（请确保在构造函数内初始化 MaterialItems 后调用）
        /// </summary>
        private void RegisterLayerPropertyTracker()
        {
            MaterialItems.CollectionChanged += (s, e) =>
            {
                if (e.NewItems != null)
                {
                    foreach (StyleMaster.Models.MaterialItem item in e.NewItems)
                    {
                        item.PropertyChanged -= Item_PropertyChanged;
                        item.PropertyChanged += Item_PropertyChanged;
                    }
                }
            };

            foreach (var item in MaterialItems)
            {
                item.PropertyChanged -= Item_PropertyChanged;
                item.PropertyChanged += Item_PropertyChanged;
            }
        }

        /// <summary>
        /// 监听模型属性改变：当 IsFrozen 改变时立即同步 CAD。
        /// </summary>
        private void Item_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "IsFrozen")
            {
                var item = sender as StyleMaster.Models.MaterialItem;
                if (item != null)
                {
                    SyncSingleLayerFrozenState(item);
                }
            }
        }

        /// <summary>
        /// 同步单个图层的冻结状态，并处理“当前层”保护逻辑。
        /// </summary>
        private void SyncSingleLayerFrozenState(StyleMaster.Models.MaterialItem item)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var db = doc.Database;
            var ed = doc.Editor;

            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                        if (lt.Has(item.LayerName))
                        {
                            var ltr = (LayerTableRecord)tr.GetObject(lt[item.LayerName], OpenMode.ForWrite);

                            // 保护逻辑：如果是当前层，禁止冻结
                            if (item.IsFrozen && db.Clayer == ltr.ObjectId)
                            {
                                ed.WriteMessage($"\n[StyleMaster] 无法冻结图层 \"{item.LayerName}\"，因为它是当前工作图层。");

                                // 临时解除监听防止死循环，回滚 UI 状态
                                item.PropertyChanged -= Item_PropertyChanged;
                                item.IsFrozen = false;
                                item.PropertyChanged += Item_PropertyChanged;
                            }
                            else
                            {
                                if (ltr.IsFrozen != item.IsFrozen)
                                {
                                    ltr.IsFrozen = item.IsFrozen;
                                }
                            }
                        }
                        tr.Commit();
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n[错误] 图层操作异常: {ex.Message}");
                    }
                }
                ed.Regen();
            }
        }

        /// <summary>
        /// 同步单个图层的冻结状态到 CAD
        /// </summary>
        private void SyncSingleLayerFrozenState(string layerName, bool isFrozen)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var db = doc.Database;
            using (doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (Autodesk.AutoCAD.DatabaseServices.LayerTable)tr.GetObject(db.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    if (lt.Has(layerName))
                    {
                        var ltr = (Autodesk.AutoCAD.DatabaseServices.LayerTableRecord)tr.GetObject(lt[layerName], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        // 执行冻结或解冻
                        ltr.IsFrozen = isFrozen;
                    }
                    tr.Commit();
                }
                doc.Editor.Regen(); // 刷新屏幕显示
            }
        }
        /// <summary>
        /// 按照指南规范提供的静态启动方法
        /// 处理单例显示逻辑
        /// </summary>
        public static void ShowTool()
        {
            if (_instance == null)
            {
                _instance = new MainControlWindow();
                _instance.Closed += (s, e) => _instance = null;
                Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(_instance);
            }
            else
            {
                _instance.Activate();
                if (_instance.WindowState == WindowState.Minimized)
                    _instance.WindowState = WindowState.Normal;
            }
        }

        /// <summary>
        /// 处理无边框窗口的标题栏拖拽逻辑
        /// </summary>
        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }

        /// <summary>
        /// 标题栏关闭按钮点击事件
        /// </summary>
        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 当鼠标在 DataGrid 上移动时触发，用于检测并启动拖拽操作
        /// </summary>
        private void DataGrid_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && sender is DataGrid dg)
            {
                // 确保是从手柄区域（TextBlock）开始拖拽
                if (e.OriginalSource is TextBlock tb && tb.Text == "⋮⋮")
                {
                    var row = UIHelpers.FindVisualParent<DataGridRow>(e.OriginalSource as DependencyObject);
                    if (row != null)
                    {
                        DragDrop.DoDragDrop(row, row.Item, DragDropEffects.Move);
                    }
                }
            }
        }
        /// <summary>
        /// 点击颜色单元格中的“...”按钮时，弹出 AutoCAD 原生色盘。
        /// 使用显式命名空间解决 DialogResult 和 Color 的引用冲突。
        /// </summary>
        private void SelectColor_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as System.Windows.Controls.Button;
            var item = btn?.DataContext as MaterialItem;
            if (item == null) return;

            Autodesk.AutoCAD.Windows.ColorDialog dlg = new Autodesk.AutoCAD.Windows.ColorDialog();
            dlg.Color = item.CadColor;

            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                item.CadColor = dlg.Color;
                // ✨ 修改：传入 item.LayerName
                item.PreviewBrush = ConvertCadColorToBrush(dlg.Color, item.LayerName);
            }
        }
        /// <summary>
        /// 当列表数据发生变动（如图层名改变）时，统一刷新所有预览色块。
        /// </summary>
        private void RefreshAllPreviews()
        {
            foreach (var item in MaterialItems)
            {
                // 先移除再添加，防止重复绑定监听
                item.PropertyChanged -= Item_PropertyChanged;
                item.PropertyChanged += Item_PropertyChanged;
            }
        }
        /// <summary>
        /// 辅助方法：将 AutoCAD 的 Color 对象转换为 WPF 画刷。
        /// 如果是“随层”，则自动从 CAD 图层表中获取该图层的真实颜色。
        /// </summary>
        private System.Windows.Media.Brush ConvertCadColorToBrush(Autodesk.AutoCAD.Colors.Color cadColor, string layerName)
        {
            // 默认回退色（若读取失败则显示灰色）
            var defaultBrush = System.Windows.Media.Brushes.Gray;

            // 情况 A：随层 (ByLayer)
            if (cadColor.IsByLayer)
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc != null && !string.IsNullOrEmpty(layerName))
                {
                    // 使用 OpenClose 事务快速读取数据库
                    using (var tr = doc.Database.TransactionManager.StartOpenCloseTransaction())
                    {
                        var lt = (LayerTable)tr.GetObject(doc.Database.LayerTableId, OpenMode.ForRead);
                        if (lt.Has(layerName))
                        {
                            var ltr = (LayerTableRecord)tr.GetObject(lt[layerName], OpenMode.ForRead);

                            // 抓取图层定义的 ColorValue (GDI 颜色)
                            System.Drawing.Color gdiColor = ltr.Color.ColorValue;
                            return new System.Windows.Media.SolidColorBrush(
                                System.Windows.Media.Color.FromRgb(gdiColor.R, gdiColor.G, gdiColor.B));
                        }
                    }
                }
                return defaultBrush;
            }

            // 情况 B：指定了具体颜色（真彩色或索引色）
            System.Drawing.Color c = cadColor.ColorValue;
            return new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(c.R, c.G, c.B));
        }
        /// <summary>
        /// 初始化测试数据。
        /// 当前已移除所有硬编码的测试项，确保程序启动时材质列表为空，等待用户手动拾取或导入。
        /// </summary>
        private void InitializeTestData()
        {
            // 已清除所有测试项代码
            RefreshPriorities();
        }
        /// <summary>
        /// 右键菜单：清空所有图层项
        /// </summary>
        private void MenuClearAll_Click(object sender, RoutedEventArgs e)
        {
            if (MaterialItems.Count > 0)
            {
                var result = System.Windows.MessageBox.Show("确定要清空所有已识别的图层映射吗？", "提示", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    MaterialItems.Clear();
                    RefreshPriorities();
                }
            }
        }

        /// <summary>
        /// 鼠标左键按下：记录起始位置和当前行对象。
        /// </summary>
        private void MainDataGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _dragStartPoint = e.GetPosition(MainDataGrid);
            _draggedItem = GetItemAtPoint(MainDataGrid, _dragStartPoint);
        }

        /// <summary>
        /// 鼠标移动：判断是否触发拖拽，并实时更新影子和指示线位置。
        /// </summary>
        private void MainDataGrid_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || _draggedItem == null || _isDraggingNow) return;

            Point currentPos = e.GetPosition(MainDataGrid);
            if (Math.Abs(currentPos.Y - _dragStartPoint.Y) > SystemParameters.MinimumVerticalDragDistance)
            {
                StartAnimatedDrag(e);
            }
        }

        /// <summary>
        /// 开始带动画的拖拽处理逻辑。
        /// 通过显式指定 System.Windows.Visibility 解决实例引用冲突。
        /// </summary>
        private void StartAnimatedDrag(MouseEventArgs e)
        {
            _isDraggingNow = true;

            // 显示并设置影子大小
            var row = MainDataGrid.ItemContainerGenerator.ContainerFromItem(_draggedItem) as DataGridRow;
            if (row != null)
            {
                DragGhost.Width = row.ActualWidth;
                DragGhost.Height = row.ActualHeight;
                // ✨ 修正：使用全称访问枚举
                DragGhost.Visibility = System.Windows.Visibility.Visible;
            }

            // 执行系统拖拽
            DragDrop.DoDragDrop(MainDataGrid, _draggedItem, DragDropEffects.Move);

            // 拖拽结束：重置 UI
            _isDraggingNow = false;
            _draggedItem = null;
            // ✨ 修正：使用全称访问枚举
            DragGhost.Visibility = System.Windows.Visibility.Collapsed;
            InsertionMarker.Visibility = System.Windows.Visibility.Collapsed;
        }

        /// <summary>
        /// 拖拽过程中实时反馈：更新影子位置和插入线。
        /// 修正了 Visibility 的静态引用问题。
        /// </summary>
        private void MainDataGrid_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Move;
            Point pos = e.GetPosition(MainDataGrid);

            // 更新影子位置 (锁定 X 轴)
            Canvas.SetTop(DragGhost, pos.Y - (DragGhost.Height / 2));
            Canvas.SetLeft(DragGhost, 0);

            // 更新插入指示线位置
            var targetItem = GetItemAtPoint(MainDataGrid, pos);
            if (targetItem != null)
            {
                var row = MainDataGrid.ItemContainerGenerator.ContainerFromItem(targetItem) as DataGridRow;
                if (row != null)
                {
                    Point rowPos = row.TranslatePoint(new Point(0, 0), MainDataGrid);
                    double markerY = (pos.Y > rowPos.Y + row.ActualHeight / 2) ? rowPos.Y + row.ActualHeight : rowPos.Y;

                    // ✨ 修正：使用全称访问枚举
                    InsertionMarker.Visibility = System.Windows.Visibility.Visible;
                    Canvas.SetTop(InsertionMarker, markerY);
                }
            }
        }

        /// <summary>
        /// 拖拽落点：执行数据交换并重置辅助 UI 状态。
        /// </summary>
        private void DataGrid_Drop(object sender, DragEventArgs e)
        {
            _isDraggingNow = false;
            // ✨ 修正：使用全称访问枚举
            DragGhost.Visibility = System.Windows.Visibility.Collapsed;
            InsertionMarker.Visibility = System.Windows.Visibility.Collapsed;

            var droppedItem = e.Data.GetData(typeof(MaterialItem)) as MaterialItem;
            var targetItem = GetItemAtPoint(MainDataGrid, e.GetPosition(MainDataGrid));

            if (droppedItem != null && targetItem != null && droppedItem != targetItem)
            {
                int oldIndex = MaterialItems.IndexOf(droppedItem);
                int newIndex = MaterialItems.IndexOf(targetItem);

                MaterialItems.Move(oldIndex, newIndex);
                RefreshPriorities();
            }
        }
        /// <summary>
        /// 辅助方法：获取坐标点下的 DataGrid 行对象。
        /// </summary>
        private MaterialItem GetItemAtPoint(DataGrid grid, Point pt)
        {
            HitTestResult hit = VisualTreeHelper.HitTest(grid, pt);
            if (hit == null) return null;

            DependencyObject parent = hit.VisualHit;
            while (parent != null && !(parent is DataGridRow))
            {
                parent = VisualTreeHelper.GetParent(parent);
            }
            return (parent as DataGridRow)?.Item as MaterialItem;
        }
        /// <summary>
        /// 重新计算并刷新集合中所有材质项的层级编号 (1, 2, 3...)
        /// </summary>
        private void RefreshPriorities()
        {
            if (MaterialItems == null) return;

            for (int i = 0; i < MaterialItems.Count; i++)
            {
                MaterialItems[i].Priority = i + 1;
            }
        }

        /// <summary>
        /// 拾取图纸范围按钮点击事件：在 CAD 中选择实体，自动提取图层并初始化为 Hatch 模式。
        /// </summary>
        private void PickLayers_Click(object sender, RoutedEventArgs e)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;

            // 引导用户进行 CAD 选择
            var promptOptions = new PromptSelectionOptions { MessageForAdding = "\n请选择图纸范围内的对象以提取图层: " };
            var selectionResult = ed.GetSelection(promptOptions);

            if (selectionResult.Status != PromptStatus.OK) return;

            var selectedLayerNames = new HashSet<string>();
            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject selObj in selectionResult.Value)
                {
                    if (selObj == null) continue;
                    var ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                    if (ent != null) selectedLayerNames.Add(ent.Layer);
                }
            }

            foreach (var layerName in selectedLayerNames)
            {
                if (MaterialItems.Any(x => x.LayerName.Equals(layerName, StringComparison.OrdinalIgnoreCase)))
                    continue;

                var newItem = new StyleMaster.Models.MaterialItem
                {
                    LayerName = layerName,
                    Priority = MaterialItems.Count + 1,
                    FillType = "Hatch",           // 固定为 Hatch 模式
                    PatternName = "SOLID",        // 默认为纯色填充

                    // ✨ 修改：移除了导致报错的 Opacity = 0

                    // ✨ 新增：初始化新增的属性
                    IsFillLayer = true,           // 默认开启材质填充
                    IsFrozen = false,             // 默认不冻结图层
                    Scale = 1.0,                  // 默认缩放比例

                    CadColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256)
                };

                // 计算图层对应的预览颜色
                newItem.PreviewBrush = ConvertCadColorToBrush(newItem.CadColor, newItem.LayerName);
                MaterialItems.Add(newItem);
            }
        }

        /// <summary>
        /// 智能匹配按钮点击事件
        /// </summary>
        private void SmartMatch_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("智能匹配功能正在开发中...");
        }

        private void RunFill_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (MaterialItems == null || MaterialItems.Count == 0) return;

            // 过滤出仅勾选了 IsFillLayer 的项
            var selectedItems = System.Linq.Enumerable.ToList(
                System.Linq.Enumerable.Where(MaterialItems, x => x.IsFillLayer)
            );

            if (selectedItems.Count == 0)
            {
                System.Windows.MessageBox.Show("请至少勾选一个需要填充的图层。");
                return;
            }

            try
            {
                // ✨ 请将 RunBatchHatch 替换为你 Service 中实际存在的方法名
                // 如果你的方法名是 RunFill，则改为：
                StyleMaster.Services.CadRenderingService.RunFill(selectedItems);

                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n[StyleMaster] 已完成勾选图层的内部填充预览。");
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show($"预览填充失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 材质选择按钮点击逻辑：弹出 Hatch 图案预览选择窗口。
        /// 已移除原有的 OpenFileDialog 图片选择逻辑，仅支持矢量 Hatch 图案。
        /// </summary>
        private void SelectMaterial_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            var item = btn?.DataContext as MaterialItem;
            if (item == null) return;

            // 统一使用 PatternSelectorWindow 进行 Hatch 图案选择
            var selector = new PatternSelectorWindow();
            selector.Owner = this;

            if (selector.ShowDialog() == true)
            {
                item.PatternName = selector.SelectedPatternName;
            }
        }

        /// <summary>
        /// 响应 DataGrid 的 Delete 按键操作
        /// </summary>
        private void MainDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                DeleteSelectedItems();
                e.Handled = true;
            }
        }

        /// <summary>
        /// 右键菜单：删除选中行点击事件
        /// </summary>
        private void MenuDelete_Click(object sender, RoutedEventArgs e)
        {
            DeleteSelectedItems();
        }

        /// <summary>
        /// 内部统一的物理删除与重排逻辑
        /// </summary>
        private void DeleteSelectedItems()
        {
            if (MainDataGrid.SelectedItems.Count > 0)
            {
                var itemsToRemove = MainDataGrid.SelectedItems.Cast<MaterialItem>().ToList();
                foreach (var item in itemsToRemove)
                {
                    MaterialItems.Remove(item);
                }
                RefreshPriorities();
            }
        }
        /// <summary>
        /// “清除填充”按钮点击事件。
        /// 遍历当前材质列表中的所有图层，并清除这些图层上已有的填充实体。
        /// </summary>
        private void ClearFills_Click(object sender, RoutedEventArgs e)
        {
            if (MaterialItems == null || MaterialItems.Count == 0) return;

            var result = System.Windows.MessageBox.Show("确定要清除当前列表中所有图层关联的填充吗？", "确认", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    // 提取列表中所有的图层名称
                    var layers = MaterialItems.Select(x => x.LayerName).ToList();
                    // 调用服务层进行批量清除
                    Services.CadRenderingService.ClearFillsOnLayers(layers);
                }
                catch (System.Exception ex)
                {
                    System.Windows.MessageBox.Show("清除填充失败: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// 响应单行刷新按钮。
        /// 强制提交 DataGrid 编辑以确保 Scale 等数值已更新，随后重新绘制该图层的 Hatch。
        /// </summary>
        private void RefreshLayer_Click(object sender, RoutedEventArgs e)
        {
            // 强制提交编辑，确保 MaterialItems 中的 Scale 属性已更新
            MainDataGrid.CommitEdit(DataGridEditingUnit.Row, true);
            MainDataGrid.Focus();

            var btn = sender as System.Windows.Controls.Button;
            var item = btn?.DataContext as MaterialItem;
            if (item == null) return;

            try
            {
                // 仅在 CAD 内执行 Hatch 填充刷新
                Services.CadRenderingService.RefreshSingleLayer(item, MaterialItems);
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show($"刷新图层 {item.LayerName} 失败: {ex.Message}");
            }
        }
        /// <summary>
        /// 导出按钮点击事件：同步导出 SVG，异步导出 PDF，并输出坐标兜底。
        /// 修改：改为 async 方法，增加了 Task.Run 异步处理打印逻辑。
        /// </summary>
        private void ExportSvg_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (MaterialItems == null || MaterialItems.Count == 0) return;

            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "SVG 文件 (*.svg)|*.svg",
                FileName = "StyleMaster_Export"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    // 1. 同步导出 SVG 并生成 DCFW 边界矩形
                    Autodesk.AutoCAD.DatabaseServices.Extents3d ext = StyleMaster.Services.CadRenderingService.ExportToSvg(MaterialItems, saveFileDialog.FileName);

                    // 2. 命令行输出坐标 (关键：供手动打印使用)
                    var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                    ed.WriteMessage("\n[StyleMaster] 导出范围已锁定 (DCFW 图层已生成)：");
                    ed.WriteMessage($"\n >> 左下角 (Min): {ext.MinPoint.X:F2}, {ext.MinPoint.Y:F2}");
                    ed.WriteMessage($"\n >> 右上角 (Max): {ext.MaxPoint.X:F2}, {ext.MaxPoint.Y:F2}");
                    ed.WriteMessage("\n[提示] 材质选区已就绪。请手动使用 Window 窗口模式打印 PDF 线稿。");

                    System.Windows.MessageBox.Show("SVG 导出成功！坐标范围已输出至命令行。");
                }
                catch (System.Exception ex)
                {
                    System.Windows.MessageBox.Show($"导出失败: {ex.Message}");
                }
            }
        }
        /// <summary>
        /// 根据列表状态一键冻结/解冻 CAD 图层
        /// </summary>
        private void ApplyLayerStates_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var lt = (Autodesk.AutoCAD.DatabaseServices.LayerTable)tr.GetObject(db.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                foreach (var item in MaterialItems)
                {
                    if (lt.Has(item.LayerName))
                    {
                        var ltr = (Autodesk.AutoCAD.DatabaseServices.LayerTableRecord)tr.GetObject(lt[item.LayerName], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        ltr.IsFrozen = item.IsFrozen; // 应用冻结状态
                    }
                }
                tr.Commit();
            }
            doc.Editor.Regen();
        }
    }
}