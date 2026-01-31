/* * 文件位置：StyleMaster/Models/MaterialItem.cs
 * 功能：材质项模型。
 * 修改说明：删除了 _isExportEnabled 字段及 IsExportEnabled 属性。
 */
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Xml.Serialization;

namespace StyleMaster.Models
{
    public class MaterialItem : INotifyPropertyChanged
    {

        private string _layerName;
        private int _priority;
        private string _fillType;
        private string _patternName;
        private double _scale = 1.0;
        private bool _isFillLayer = true;
        private bool _isFrozen = false;
        private int _colorIndex = 256;

        private Autodesk.AutoCAD.Colors.Color _cadColor;
        private System.Windows.Media.Brush _previewBrush;

        public string LayerName { get => _layerName; set { _layerName = value; OnPropertyChanged(); } }
        public int Priority { get => _priority; set { _priority = value; OnPropertyChanged(); } }
        public string FillType { get => _fillType; set { _fillType = value; OnPropertyChanged(); } }
        public string PatternName { get => _patternName; set { _patternName = value; OnPropertyChanged(); } }
        public double Scale { get => _scale; set { _scale = value; OnPropertyChanged(); } }
        public bool IsFillLayer { get => _isFillLayer; set { _isFillLayer = value; OnPropertyChanged(); } }
        public bool IsFrozen { get => _isFrozen; set { _isFrozen = value; OnPropertyChanged(); } }

        public int ColorIndex { get => _colorIndex; set { _colorIndex = value; OnPropertyChanged(); } }

        [XmlIgnore]
        public Autodesk.AutoCAD.Colors.Color CadColor
        {
            get => _cadColor;
            set
            {
                _cadColor = value;
                if (value != null) _colorIndex = value.ColorIndex;
                OnPropertyChanged();
            }
        }

        [XmlIgnore]
        public System.Windows.Media.Brush PreviewBrush { get => _previewBrush; set { _previewBrush = value; OnPropertyChanged(); } }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}