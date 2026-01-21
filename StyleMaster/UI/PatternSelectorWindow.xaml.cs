using StyleMaster.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Input;

namespace StyleMaster.UI
{
    public partial class PatternSelectorWindow : Window
    {
        private List<PatternItem> _allPatterns = new List<PatternItem>();
        public string SelectedPatternName { get; private set; }

        public PatternSelectorWindow()
        {
            InitializeComponent();
            LoadPatterns();
        }

        /// <summary>
        /// 扫描 Patterns 文件夹及其隐藏的 .hatch_thumbs 文件夹
        /// </summary>
        private void LoadPatterns()
        {
            try
            {
                string rootDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string patternsPath = Path.Combine(rootDir, "Resources", "Patterns");
                string thumbsPath = Path.Combine(patternsPath, ".hatch_thumbs");

                if (!Directory.Exists(patternsPath)) Directory.CreateDirectory(patternsPath);
                if (!Directory.Exists(thumbsPath)) Directory.CreateDirectory(thumbsPath);

                // 读取收藏配置 (此处可扩展为 JSON 读取)
                var favorites = new HashSet<string>();

                var files = Directory.GetFiles(patternsPath, "*.pat");
                foreach (var file in files)
                {
                    string name = Path.GetFileNameWithoutExtension(file);
                    string thumb = Path.Combine(thumbsPath, name + ".png");

                    _allPatterns.Add(new PatternItem
                    {
                        Name = name,
                        ThumbnailPath = File.Exists(thumb) ? thumb : null, // ✨ 改为 null
                        IsFavorite = favorites.Contains(name)
                    });
                }

                RefreshDisplay();
            }
            catch (System.Exception ex) // ✨ 显式指定 System 命名空间
            {
                MessageBox.Show("加载图案库失败: " + ex.Message);
            }
        }

        private void RefreshDisplay()
        {
            var filtered = _allPatterns.AsEnumerable();

            // 搜索过滤
            if (!string.IsNullOrEmpty(SearchBox.Text))
                filtered = filtered.Where(p => p.Name.ToLower().Contains(SearchBox.Text.ToLower()));

            // 收藏过滤
            // if (FilterCheckbox.IsChecked == true) filtered = filtered.Where(p => p.IsFavorite);

            PatternList.ItemsSource = filtered.OrderByDescending(p => p.IsFavorite).ThenBy(p => p.Name).ToList();
        }

        private void Item_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.DataContext is PatternItem item)
            {
                SelectedPatternName = item.Name;
                this.DialogResult = true;
                this.Close();
            }
        }

        private void Star_Click(object sender, RoutedEventArgs e)
        {
            // 此处处理收藏状态持久化逻辑
            RefreshDisplay();
        }

        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e) => this.DragMove();
        private void Close_Click(object sender, RoutedEventArgs e) => this.Close();
        private void SearchBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e) => RefreshDisplay();
        private void Filter_Changed(object sender, RoutedEventArgs e) => RefreshDisplay();
    }
}