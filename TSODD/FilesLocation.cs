using System.IO;
using System.Reflection;

namespace TSODD
{
    internal static class FilesLocation
    {
        private static string dllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        public static DirectoryInfo SettingsFolder = Directory.CreateDirectory(System.IO.Path.Combine(dllPath, "Support"));
        public static string JsonTableNamesSignsPath = Path.GetFullPath(Path.Combine(SettingsFolder.FullName, "NamesSigns.json"));
        public static string JsonTableHeaderSignsPath = Path.GetFullPath(Path.Combine(SettingsFolder.FullName, "TableHeaderSigns.json"));
        public static string JsonTableHeaderMarksPath = Path.GetFullPath(Path.Combine(SettingsFolder.FullName, "TableHeadersMarks.json"));
        public static string JsonOptionsPath = Path.GetFullPath(Path.Combine(SettingsFolder.FullName, "UserOptions.json"));
        public static string ExcelTableSignsPath = Path.GetFullPath(Path.Combine(SettingsFolder.FullName, "Дорожные знаки.xlsx"));

        public static string dwgBlocksPath = Path.Combine(dllPath, "Support", "blocks.dwg");
        public static string dwgTemplatePath = Path.Combine(dllPath, "Support", "templates.dwg");
        public static string linPath = Path.Combine(dllPath, "Support", "LineTypes.lin");
        public static string separatedLinPath = Path.Combine(dllPath, "Support", "SeparatedLineTypes.lin");


    }
}
