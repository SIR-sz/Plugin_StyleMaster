using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace StyleMaster.Models
{
    /// <summary>
    /// 填充图案显示项模型
    /// </summary>
    public class PatternItem : INotifyPropertyChanged
    {
        private bool _isFavorite;

        /// <summary>
        /// 图案名称（对应 .pat 文件名）
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 缩略图完整路径
        /// </summary>
        public string ThumbnailPath { get; set; }

        /// <summary>
        /// 是否已收藏
        /// </summary>
        public bool IsFavorite
        {
            get => _isFavorite;
            set { _isFavorite = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}