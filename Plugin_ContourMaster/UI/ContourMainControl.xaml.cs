using Plugin_ContourMaster.Models;
using Plugin_ContourMaster.Services;
using System.Windows;
using System.Windows.Controls;

namespace Plugin_ContourMaster.UI
{
    public partial class ContourMainControl : UserControl
    {
        // ✨ 定义的是 Settings（大写）
        public ContourSettings Settings { get; set; } = new ContourSettings();

        public ContourMainControl()
        {
            InitializeComponent();
            this.DataContext = Settings;
        }

        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            RunEngineAction(engine => engine.ProcessContour());
        }

        private void BtnImage_Click(object sender, RoutedEventArgs e)
        {
            // ✨ 统一使用 Settings 属性
            RunEngineAction(engine => engine.ProcessReferencedEntity(false));
        }

        private void BtnPdf_Click(object sender, RoutedEventArgs e)
        {
            // ✨ 统一使用 Settings 属性
            RunEngineAction(engine => engine.ProcessReferencedEntity(true));
        }

        private void RunEngineAction(System.Action<ContourEngine> action)
        {
            try
            {
                // ✨ 统一使用 Settings
                var engine = new ContourEngine(Settings);
                action.Invoke(engine);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"操作失败: {ex.Message}", "轮廓提取", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}