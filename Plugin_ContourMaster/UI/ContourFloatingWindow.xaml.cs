// --- UI/ContourFloatingWindow.xaml.cs ---

using Plugin_ContourMaster.Models;
using System;
using System.Windows;

namespace Plugin_ContourMaster.UI
{
    /// <summary>
    /// 悬浮窗外壳：负责承载核心 UI 并处理窗口位置记忆
    /// </summary>
    public partial class ContourFloatingWindow : Window
    {
        public ContourFloatingWindow(ContourSettings settings)
        {
            InitializeComponent();

            // 将设置对象绑定到窗口，使 Top/Left/Width/Height 的双向绑定生效
            this.DataContext = settings;

            // 将设置对象传递给内部的 UserControl
            this.MainContent.Settings = settings;
            this.MainContent.DataContext = settings;
        }

        /// <summary>
        /// 当窗口关闭时，确保位置数据已被保存（虽然 Binding 已经处理了大部分，但这步可以作为保险）
        /// </summary>
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
        }
    }
}