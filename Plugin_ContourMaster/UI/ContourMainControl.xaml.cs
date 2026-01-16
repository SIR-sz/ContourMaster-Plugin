using Plugin_ContourMaster.Models;
using Plugin_ContourMaster.Services;
using System.Windows;
using System.Windows.Controls;

namespace Plugin_ContourMaster.UI
{
    /// <summary>
    /// ContourMainControl.xaml 的交互逻辑
    /// </summary>
    public partial class ContourMainControl : UserControl
    {
        // 实例化设置模型
        public ContourSettings Settings { get; set; } = new ContourSettings();

        public ContourMainControl()
        {
            InitializeComponent();

            // 将 DataContext 绑定到设置对象，实现 Slider 和 TextBox 的双向同步
            this.DataContext = Settings;
        }

        /// <summary>
        /// 点击按钮触发核心算法
        /// </summary>
        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 调用服务层的引擎
                var engine = new ContourEngine(Settings);
                engine.ProcessContour();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"启动失败: {ex.Message}", "像素轮廓专家", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}