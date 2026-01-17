using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace Plugin_ContourMaster.Models
{
    [Obfuscation(Feature = "renaming", Exclude = true)]
    public class ContourSettings : INotifyPropertyChanged
    {
        private double _threshold = 128.0;
        private double _simplifyTolerance = 0.5;
        private int _smoothLevel = 2;
        private string _layerName = "LK_XS"; // ✨ 修改：默认图层改为 LK_XS
        private int _precisionLevel = 5;     // ✨ 采样精度等级 (1-10)，默认 5
                                             // 在 ContourSettings 类中添加
        private bool _isOcrMode = false;
        public bool IsOcrMode
        {
            get => _isOcrMode;
            set { _isOcrMode = value; OnPropertyChanged(); }
        }

        public double Threshold
        {
            get => _threshold;
            set { _threshold = value; OnPropertyChanged(); }
        }

        public double SimplifyTolerance
        {
            get => _simplifyTolerance;
            set { _simplifyTolerance = value; OnPropertyChanged(); }
        }

        public int SmoothLevel
        {
            get => _smoothLevel;
            set { _smoothLevel = value; OnPropertyChanged(); }
        }

        public string LayerName
        {
            get => _layerName;
            set { _layerName = value; OnPropertyChanged(); }
        }

        public int PrecisionLevel
        {
            get => _precisionLevel;
            set { _precisionLevel = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}