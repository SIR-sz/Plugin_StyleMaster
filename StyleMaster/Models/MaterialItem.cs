using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;

namespace StyleMaster.Models
{
    /// <summary>
    /// 图层材质模型类 - 最终修正版
    /// 修复：统一私有变量为 _fillType，公开属性为 FillType，彻底解决 CS0103 错误。
    /// </summary>
    public class MaterialItem : INotifyPropertyChanged
    {
        private string _layerName;
        private string _patternName;
        private bool _isFillLayer = true;
        private Autodesk.AutoCAD.Colors.Color _cadColor;
        private double _scale = 1.0;
        private int _opacity = 100;
        private int _priority = 0;
        private string _fillType = "Hatch"; // 统一使用 _fillType
        private Brush _previewBrush = Brushes.Gray;

        public string LayerName
        {
            get => _layerName;
            set { _layerName = value; OnPropertyChanged(); }
        }

        public string PatternName
        {
            get => _patternName;
            set { _patternName = value; OnPropertyChanged(); }
        }

        public bool IsFillLayer
        {
            get => _isFillLayer;
            set { _isFillLayer = value; OnPropertyChanged(); }
        }

        public Autodesk.AutoCAD.Colors.Color CadColor
        {
            get => _cadColor;
            set { _cadColor = value; OnPropertyChanged(); }
        }

        public double Scale
        {
            get => _scale;
            set { _scale = value; OnPropertyChanged(); }
        }

        public int Opacity
        {
            get => _opacity;
            set { _opacity = value; OnPropertyChanged(); }
        }

        public int Priority
        {
            get => _priority;
            set { _priority = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// 填充类型 (对应 UI 逻辑中的 FillType)
        /// </summary>
        public string FillType
        {
            get => _fillType;
            set { _fillType = value; OnPropertyChanged(); }
        }

        public Brush PreviewBrush
        {
            get => _previewBrush;
            set { _previewBrush = value; OnPropertyChanged(); }
        }

        public string ColorDescription => CadColor != null ? CadColor.ToString() : "未设置";

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}