using StyleMaster.Models;
using System;
using System.Collections.ObjectModel;
using System.Linq;
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
        /// 遍历集合，根据当前 UI 顺序重新分配层级数字 (Priority)
        /// </summary>
        private void RefreshPriorities()
        {
            for (int i = 0; i < MaterialItems.Count; i++)
            {
                MaterialItems[i].Priority = i + 1;
            }
        }
        /// <summary>
        /// 拾取图纸范围按钮点击事件：自动提取图层
        /// </summary>
        private void PickLayers_Click(object sender, RoutedEventArgs e)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            // 1. 暂时隐藏窗口
            this.Hide();

            try
            {
                // 2. 提示用户选择物体
                var psr = ed.GetSelection();
                if (psr.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                {
                    using (var tr = doc.TransactionManager.StartTransaction())
                    {
                        var newLayers = new System.Collections.Generic.HashSet<string>();
                        foreach (var id in psr.Value.GetObjectIds())
                        {
                            var ent = tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                            if (ent != null)
                            {
                                newLayers.Add(ent.Layer);
                            }
                        }

                        // 3. 将新图层合并到列表（去重）
                        foreach (var layer in newLayers)
                        {
                            if (!MaterialItems.Any(m => m.LayerName == layer))
                            {
                                MaterialItems.Add(new MaterialItem { LayerName = layer });
                            }
                        }
                        RefreshPriorities();
                        ed.WriteMessage($"\n[StyleMaster] 已成功识别并添加 {newLayers.Count} 个图层。");
                    }
                }
            }
            finally
            {
                // 4. 恢复窗口显示
                this.Show();
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
        /// 材质选择按钮点击逻辑（分模式处理）
        /// </summary>
        private void SelectMaterial_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            var item = btn?.DataContext as MaterialItem;
            if (item == null) return;

            if (item.FillMode == FillType.Hatch)
            {
                // TODO: 弹出你要求的“简化版图案预览窗口”
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("即将弹出 Hatch 图案预览库...");
            }
            else
            {
                // 弹出图片选择对话框
                var dialog = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "图片材质|*.png;*.jpg;*.jpeg;*.bmp",
                    Title = "选择材质图片"
                };
                if (dialog.ShowDialog() == true)
                {
                    item.PatternName = dialog.FileName;
                }
            }
        }
    }
}