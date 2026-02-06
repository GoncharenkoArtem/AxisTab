using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace AxisTab
{

    public static class WindowAcadActivity
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        private static uint? _autocadProcessId = null;

        /// <summary>
        /// Проверяет, является ли AutoCAD активным приложением (в фокусе)
        /// </summary>
        public static bool IsAutoCADActive()
        {
            try
            {
                var foregroundWindow = GetForegroundWindow();
                if (foregroundWindow == IntPtr.Zero)
                    return false;

                uint windowProcessId;
                GetWindowThreadProcessId(foregroundWindow, out windowProcessId);

                // Получаем ID процесса AutoCAD один раз
                if (!_autocadProcessId.HasValue)
                {
                    _autocadProcessId = (uint)Process.GetCurrentProcess().Id;
                }

                // Проверяем, совпадает ли процесс активного окна с процессом AutoCAD
                return windowProcessId == _autocadProcessId.Value;
            }
            catch
            {
                return false;
            }
        }
    }
}
