using System.IO;
using System.Reflection;

namespace AxisTAb
{
    internal static class FilesLocation
    {
        private static string dllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        public static DirectoryInfo SettingsFolder = Directory.CreateDirectory(System.IO.Path.Combine(dllPath, "Support"));
        public static string JsonOptionsPath = Path.GetFullPath(Path.Combine(SettingsFolder.FullName, "UserOptions.json"));

    }
}
