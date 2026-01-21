using Autodesk.AutoCAD.Colors;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;

namespace StyleMaster.Models
{
    /// <summary>
    /// 填充类型枚举
    /// </summary>
    public enum FillType
    {
        Hatch,
        Image
    }
    /// <summary>
    /// 材质映射项模型。
    /// 通过显式指定命名空间解决 Autodesk.AutoCAD.Colors.Color 与 System.Windows.Media.Color 的冲突。
    /// </summary>
    public class MaterialItem : INotifyPropertyChanged
    {
        private int _priority;
        private string _layerName;
        private FillType _fillMode;
        private string _patternName;
        private double _scale = 1.0;
        private double _opacity = 0;
        private Autodesk.AutoCAD.Colors.Color _cadColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 256);

        // ✨ 确保这里只定义了一次私有字段
        private Brush _previewBrush = Brushes.Gray;

        public int Priority { get => _priority; set { _priority = value; OnPropertyChanged(); } }
        public string LayerName { get => _layerName; set { _layerName = value; OnPropertyChanged(); } }
        public FillType FillMode { get => _fillMode; set { _fillMode = value; OnPropertyChanged(); } }
        public string PatternName { get => _patternName; set { _patternName = value; OnPropertyChanged(); } }
        public double Scale { get => _scale; set { _scale = value; OnPropertyChanged(); } }
        public double Opacity { get => _opacity; set { _opacity = value; OnPropertyChanged(); } }

        public Autodesk.AutoCAD.Colors.Color CadColor
        {
            get => _cadColor;
            set
            {
                _cadColor = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ColorDescription));
            }
        }

        public string ColorDescription => CadColor.IsByLayer ? "随层" : (CadColor.HasColorName ? CadColor.ColorName : CadColor.ToString());

        // ✨ 属性定义
        public Brush PreviewBrush
        {
            get => _previewBrush;
            set { _previewBrush = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}