using System.ComponentModel;
using System.Runtime.CompilerServices;

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
    /// 材质映射项模型
    /// </summary>
    public class MaterialItem : INotifyPropertyChanged
    {
        private int _priority;
        private string _layerName;
        private FillType _fillMode = FillType.Hatch;
        private string _patternName = "SOLID";
        private string _foreColor = "随层";
        private string _backColor = "None";
        private double _scale = 1.0;
        private double _opacity = 100;

        public int Priority
        {
            get => _priority;
            set { _priority = value; OnPropertyChanged(); }
        }

        public string LayerName
        {
            get => _layerName;
            set { _layerName = value; OnPropertyChanged(); }
        }

        public FillType FillMode
        {
            get => _fillMode;
            set { _fillMode = value; OnPropertyChanged(); }
        }

        public string PatternName
        {
            get => _patternName;
            set { _patternName = value; OnPropertyChanged(); }
        }

        public string ForeColor
        {
            get => _foreColor;
            set { _foreColor = value; OnPropertyChanged(); }
        }

        public string BackColor
        {
            get => _backColor;
            set { _backColor = value; OnPropertyChanged(); }
        }

        public double Scale
        {
            get => _scale;
            set { _scale = value; OnPropertyChanged(); }
        }

        public double Opacity
        {
            get => _opacity;
            set { _opacity = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}