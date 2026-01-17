using Plugin_ContourMaster.Models;
using Plugin_ContourMaster.Services;
using System.Windows;
using System.Windows.Controls;

namespace Plugin_ContourMaster.UI
{
    public partial class ContourMainControl : UserControl
    {
        // ✨ 这里定义的是 Settings
        public ContourSettings Settings { get; set; } = new ContourSettings();

        public ContourMainControl()
        {
            InitializeComponent();
            this.DataContext = Settings;
        }

        // 按钮 1：开始提取轮廓 (CAD 选区)
        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            RunEngineAction(engine => engine.ProcessContour());
        }

        // 按钮 2：识别图片
        private void BtnImage_Click(object sender, RoutedEventArgs e)
        {
            RunEngineAction(engine => engine.ProcessImageContour());
        }

        // 按钮 3：识别 PDF
        private void BtnPdf_Click(object sender, RoutedEventArgs e)
        {
            RunEngineAction(engine => engine.ProcessPdfContour());
        }

        private void RunEngineAction(System.Action<ContourEngine> action)
        {
            try
            {
                // ✨ 统一使用定义的 Settings 属性
                var engine = new ContourEngine(Settings);
                action.Invoke(engine);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"操作失败: {ex.Message}", "像素轮廓专家", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}