using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace AxisTAb
{
    internal static class JsonReader
    {
        // метод сохранения в json
        public static void SaveToJson<T>(T obj, string path)
        {
            string json = JsonSerializer.Serialize(obj, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }

        // метод загрузки из json
        public static T LoadFromJson<T>(string path)
        {
            string json = File.ReadAllText(path);
            var obj = JsonSerializer.Deserialize<T>(json);
            return obj;
        }
    }
}
