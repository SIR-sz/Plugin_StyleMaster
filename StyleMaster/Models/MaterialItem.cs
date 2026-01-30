/* * 文件位置：StyleMaster/Models/MaterialItem.cs
 * 功能：材质项模型。
 * 修改说明：通过显式指定全限定名（Autodesk.AutoCAD.Colors.Color）解决了 AutoCAD 与 WPF 命名空间中 Color 类的引用冲突。
 */
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace StyleMaster.Models
{
    public class MaterialItem : INotifyPropertyChanged
    {
        private bool _isExportEnabled = true;
        private string _layerName;
        private int _priority;
        private string _fillType;
        private string _patternName;
        private double _scale = 1.0;
        private Autodesk.AutoCAD.Colors.Color _cadColor; // 修改点：使用全限定名避免冲突
        private System.Windows.Media.Brush _previewBrush;
        private bool _isFillLayer = true;
        private bool _isFrozen = false;

        /// <summary>
        /// 是否参与导出/打印
        /// </summary>
        public bool IsExportEnabled
        {
            get => _isExportEnabled;
            set { _isExportEnabled = value; OnPropertyChanged(); }
        }

        public string LayerName
        {
            get => _layerName;
            set { _layerName = value; OnPropertyChanged(); }
        }

        public int Priority
        {
            get => _priority;
            set { _priority = value; OnPropertyChanged(); }
        }

        public string FillType
        {
            get => _fillType;
            set { _fillType = value; OnPropertyChanged(); }
        }

        public string PatternName
        {
            get => _patternName;
            set { _patternName = value; OnPropertyChanged(); }
        }

        public double Scale
        {
            get => _scale;
            set { _scale = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// AutoCAD 颜色对象
        /// </summary>
        public Autodesk.AutoCAD.Colors.Color CadColor // 修改点：使用全限定名避免冲突
        {
            get => _cadColor;
            set { _cadColor = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// WPF 预览画刷
        /// </summary>
        public System.Windows.Media.Brush PreviewBrush
        {
            get => _previewBrush;
            set { _previewBrush = value; OnPropertyChanged(); }
        }

        public bool IsFillLayer
        {
            get => _isFillLayer;
            set { _isFillLayer = value; OnPropertyChanged(); }
        }

        public bool IsFrozen
        {
            get => _isFrozen;
            set { _isFrozen = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}