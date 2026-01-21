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
        // ✨ 修正：使用全称避免歧义
        private Autodesk.AutoCAD.Colors.Color _cadColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256);
        private System.Windows.Media.Brush _previewBrush = System.Windows.Media.Brushes.Gray;

        public int Priority { get => _priority; set { _priority = value; OnPropertyChanged(); } }
        public string LayerName { get => _layerName; set { _layerName = value; OnPropertyChanged(); } }
        public FillType FillMode { get => _fillMode; set { _fillMode = value; OnPropertyChanged(); } }
        public string PatternName { get => _patternName; set { _patternName = value; OnPropertyChanged(); } }
        public double Scale { get => _scale; set { _scale = value; OnPropertyChanged(); } }
        public double Opacity { get => _opacity; set { _opacity = value; OnPropertyChanged(); } }

        /// <summary>
        /// 存储 AutoCAD 原生颜色对象（显式指定命名空间）
        /// </summary>
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

        /// <summary>
        /// 用于 UI 显示的颜色文字描述
        /// </summary>
        public string ColorDescription => CadColor.IsByLayer ? "随层" : (CadColor.HasColorName ? CadColor.ColorName : CadColor.ToString());

        /// <summary>
        /// 用于 UI 预览方块的 Brush 对象
        /// </summary>
        public System.Windows.Media.Brush PreviewBrush { get => _previewBrush; set { _previewBrush = value; OnPropertyChanged(); } }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}