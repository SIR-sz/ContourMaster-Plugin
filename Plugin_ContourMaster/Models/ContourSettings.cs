using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace Plugin_ContourMaster.Models
{
    // ✨ 新增：OCR 引擎类型枚举
    public enum OcrEngineType
    {
        Tesseract,
        Baidu
    }

    [Obfuscation(Feature = "renaming", Exclude = true)]
    public class ContourSettings : INotifyPropertyChanged
    {
        private double _threshold = 128.0;
        private double _simplifyTolerance = 0.5;
        private int _smoothLevel = 2;
        private string _layerName = "LK_XS";
        private int _precisionLevel = 5;
        private bool _isOcrMode = false;

        // ✨ 新增：百度 OCR 相关字段
        private OcrEngineType _selectedOcrEngine = OcrEngineType.Tesseract;
        private string _baiduApiKey = "";
        private string _baiduSecretKey = "";

        public bool IsOcrMode
        {
            get => _isOcrMode;
            set { _isOcrMode = value; OnPropertyChanged(); }
        }

        // ✨ 新增：OCR 引擎选择属性
        public OcrEngineType SelectedOcrEngine
        {
            get => _selectedOcrEngine;
            set { _selectedOcrEngine = value; OnPropertyChanged(); }
        }

        // ✨ 新增：百度 API Key
        public string BaiduApiKey
        {
            get => _baiduApiKey;
            set { _baiduApiKey = value; OnPropertyChanged(); }
        }

        // ✨ 新增：百度 Secret Key
        public string BaiduSecretKey
        {
            get => _baiduSecretKey;
            set { _baiduSecretKey = value; OnPropertyChanged(); }
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