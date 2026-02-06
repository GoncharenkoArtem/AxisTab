
namespace AxisTab
{
    public class Options
    {
   
        // ПК
        public double PKLineLength { get; set; } = 1;
        public string PKTextStyle { get; set; } = "";
        public double PKTextHeight { get; set; } = 2.5;

        // Бездействие
        public bool Inactivity { get; set; } = true;
        public double InactivityTimeSpan { get; set; } = 4.5;
        public double InactivityFullTime { get; set; } = 5;
    }
}

