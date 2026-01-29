using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;

namespace StyleMaster.Models
{
    public class MaterialItem : INotifyPropertyChanged
    {
        private string _layerName;
        private string _patternName;
        private bool _isFillLayer = true;
        private bool _isFrozen = false; // ✨ 新增：冻结属性
        private Autodesk.AutoCAD.Colors.Color _cadColor;
        private double _scale = 1.0;
        private int _priority = 0;
        private string _fillType = "Hatch";
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

        /// <summary>
        /// 是否冻结图层
        /// </summary>
        public bool IsFrozen
        {
            get => _isFrozen;
            set
            {
                if (_isFrozen != value)
                {
                    _isFrozen = value;
                    OnPropertyChanged(); // ✨ 必须触发此通知，UI 勾选才会生效
                }
            }
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