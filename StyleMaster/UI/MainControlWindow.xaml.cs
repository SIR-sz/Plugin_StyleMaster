using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using StyleMaster.Models;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace StyleMaster.UI
{
    /// <summary>
    /// MainControlWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainControlWindow : Window
    {
        private static MainControlWindow _instance;

        // 数据源集合
        public ObservableCollection<MaterialItem> MaterialItems { get; set; }

        /// <summary>
        /// 构造函数：初始化数据集合并设置绑定上下文
        /// </summary>
        public MainControlWindow()
        {
            InitializeComponent();

            // 初始化集合并填充假数据用于原型测试
            MaterialItems = new ObservableCollection<MaterialItem>();
            InitializeTestData();

            // 设置 DataContext 方便 XAML 绑定
            this.DataContext = this;
            this.MainDataGrid.ItemsSource = MaterialItems;
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
        /// 填充初期测试数据，用于验证 UI 效果
        /// 修复说明：将字符串直接赋值改为使用 FillType 枚举值
        /// </summary>
        private void InitializeTestData()
        {
            MaterialItems.Add(new MaterialItem { Priority = 1, LayerName = "AR-投影", FillMode = FillType.Hatch, PatternName = "SOLID", ForeColor = "黑色", BackColor = "None", Opacity = 50 });
            MaterialItems.Add(new MaterialItem { Priority = 2, LayerName = "AR-草地", FillMode = FillType.Hatch, PatternName = "GRASS", ForeColor = "深绿", BackColor = "浅绿", Scale = 1.0 });
            MaterialItems.Add(new MaterialItem { Priority = 3, LayerName = "AR-铺装", FillMode = FillType.Image, PatternName = "石材01.png", Scale = 2.0 });
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
        /// 核心逻辑：处理拖拽完成后的数据交换与重排
        /// </summary>
        private void DataGrid_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(MaterialItem)))
            {
                var droppedItem = e.Data.GetData(typeof(MaterialItem)) as MaterialItem;
                var targetRow = UIHelpers.FindVisualParent<DataGridRow>(e.OriginalSource as DependencyObject);

                if (droppedItem != null && targetRow != null)
                {
                    var targetItem = targetRow.Item as MaterialItem;
                    int oldIndex = MaterialItems.IndexOf(droppedItem);
                    int newIndex = MaterialItems.IndexOf(targetItem);

                    if (oldIndex != newIndex && oldIndex != -1 && newIndex != -1)
                    {
                        // 在集合中移动元素
                        MaterialItems.Move(oldIndex, newIndex);
                        // 重新计算并更新所有行的 Priority 数字
                        RefreshPriorities();
                    }
                }
            }
        }
        /// <summary>
        /// [辅助逻辑] 重新计算并刷新集合中所有材质项的层级编号 (1, 2, 3...)
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
        /// 拾取图纸范围按钮点击事件：自动提取选中多段线的图层
        /// </summary>
        private void PickLayers_Click(object sender, RoutedEventArgs e)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            // 1. 暂时隐藏窗口，避免挡住 CAD 操作界面
            this.Hide();

            try
            {
                // 2. 设置选择过滤器：只选择多段线
                TypedValue[] tvs = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Operator, "<OR"),
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                    new TypedValue((int)DxfCode.Start, "POLYLINE"),
                    new TypedValue((int)DxfCode.Operator, "OR>")
                };
                SelectionFilter filter = new SelectionFilter(tvs);

                // 3. 提示用户进行框选
                PromptSelectionResult psr = ed.GetSelection(filter);

                if (psr.Status == PromptStatus.OK)
                {
                    using (var tr = doc.TransactionManager.StartTransaction())
                    {
                        var pickedLayers = new System.Collections.Generic.HashSet<string>();

                        foreach (ObjectId id in psr.Value.GetObjectIds())
                        {
                            var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (ent != null)
                            {
                                // 提取图层名
                                pickedLayers.Add(ent.Layer);
                            }
                        }

                        // 4. 将新图层合并到 DataGrid 的数据源中
                        int addCount = 0;
                        foreach (var layerName in pickedLayers)
                        {
                            // 检查是否已存在，避免重复添加
                            if (!MaterialItems.Any(x => x.LayerName == layerName))
                            {
                                MaterialItems.Add(new MaterialItem
                                {
                                    LayerName = layerName,
                                    FillMode = FillType.Hatch,
                                    PatternName = "SOLID"
                                });
                                addCount++;
                            }
                        }

                        // 5. 刷新界面层级数字
                        RefreshPriorities();

                        ed.WriteMessage($"\n[StyleMaster] 拾取完成：识别到 {pickedLayers.Count} 个图层，其中新增 {addCount} 个。");
                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[错误] 拾取失败: {ex.Message}");
            }
            finally
            {
                // 6. 无论操作是否成功或取消，都必须重新显示窗口
                this.Show();
                this.Activate(); // 确保窗口回到最前
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
                // 注意：由于正在遍历，需转为 List 处理
                var itemsToRemove = MainDataGrid.SelectedItems.Cast<MaterialItem>().ToList();
                foreach (var item in itemsToRemove)
                {
                    MaterialItems.Remove(item);
                }
                RefreshPriorities();
            }
        }

        /// <summary>
        /// 材质选择按钮点击逻辑：针对 Hatch 模式弹出预览选择窗口
        /// </summary>
        private void SelectMaterial_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as System.Windows.Controls.Button;
            var item = btn?.DataContext as MaterialItem;
            if (item == null) return;

            if (item.FillMode == FillType.Hatch)
            {
                // 弹出自定义图案预览窗口
                var selector = new PatternSelectorWindow();
                selector.Owner = this; // 设置所有者，确保居中显示且任务栏一致

                if (selector.ShowDialog() == true)
                {
                    // 选中确认后立即关闭，并更新模型数据
                    item.PatternName = selector.SelectedPatternName;
                }
            }
            else
            {
                // Image 模式保持原有的文件选择逻辑
                var dialog = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "图片材质|*.png;*.jpg;*.jpeg;*.bmp",
                    InitialDirectory = System.IO.Path.Combine(
                        System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location),
                        "Resources", "Materials")
                };
                if (dialog.ShowDialog() == true)
                {
                    item.PatternName = System.IO.Path.GetFileName(dialog.FileName);
                }
            }
        }
        /// <summary>
        /// “智能匹配”按钮点击事件（暂留接口，后期实现算法）
        /// </summary>
        private void SmartMatch_Click(object sender, RoutedEventArgs e)
        {
            // TODO: 调用 MaterialService 执行模糊匹配逻辑
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n[StyleMaster] 正在开发中：将根据图层关键字自动匹配库中材质...");
        }

        /// <summary>
        /// “一键填充”按钮点击事件
        /// </summary>
        private void RunFill_Click(object sender, RoutedEventArgs e)
        {
            if (MaterialItems == null || MaterialItems.Count == 0)
            {
                System.Windows.MessageBox.Show("列表为空，请先拾取图层。");
                return;
            }

            try
            {
                // 调用渲染服务执行填充
                StyleMaster.Services.CadRenderingService.ExecuteFill(MaterialItems);
            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("填充失败: " + ex.Message);
            }
        }
    }
}