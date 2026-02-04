
namespace AxisTab
{
    public class Options
    {
        // блоки
        public bool BlocksAnnotativeState { get; set; } = false;
        public double BlockBindRadius { get; set; } = 5;
        public string BlockNameTextStyle { get; set; } = "";
        public double BlockNameTextHeight { get; set; } = 2.5;

        // ПК
        public double PKLineLength { get; set; } = 1;
        public string PKTextStyle { get; set; } = "";
        public double PKTextHeight { get; set; } = 2.5;

        // Разметка линии
        public string LineTypeTextStyle { get; set; } = "";
        public double LineTypeTextHeight { get; set; } = 2.5;

        // Мультивыноска 
        public string MleaderStyle { get; set; } = "";
        public double MleaderTextHeight { get; set; } = 2.5;

    }
}

