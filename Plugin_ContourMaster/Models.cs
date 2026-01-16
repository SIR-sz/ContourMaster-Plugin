using System.Reflection;

namespace Plugin_ContourMaster
{
    [Obfuscation(Feature = "renaming", Exclude = true)]
    public class ContourSettings
    {
        public double Threshold { get; set; } = 128.0; // 灰度阈值
        public bool IsInverse { get; set; } = false;   // 是否反向选择
        public double SimplifyTolerance { get; set; } = 0.5; // 曲线简化公差
        public string LayerName { get; set; } = "Contour_Result"; // 结果图层

    }

}