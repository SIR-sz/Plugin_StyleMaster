using System.Windows;
using System.Windows.Media;

namespace StyleMaster.UI
{
    /// <summary>
    /// UI 树辅助工具类
    /// </summary>
    public static class UIHelpers
    {
        /// <summary>
        /// 向上递归查找指定类型的视觉父级元素
        /// </summary>
        public static T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);
            if (parentObject == null) return null;
            if (parentObject is T parent) return parent;
            return FindVisualParent<T>(parentObject);
        }
    }
}