using System.IO;
using System.Text.Json;

namespace AutoSaveAddIn
{
    public class JsonSerializationUtility
    {
        private static JsonSerializerOptions _options = new JsonSerializerOptions()
        {
            WriteIndented = true,
        };

        public static void SerializeToJson<T>(T obj, string filePath)
        {
            var json = JsonSerializer.Serialize(obj, _options);

            var fileInfo = new FileInfo(filePath);
            if (!fileInfo.Directory.Exists)
            {
                Directory.CreateDirectory(fileInfo.DirectoryName);
            }

            File.WriteAllText(filePath, json);
        }

        public static T DeserializeFromJson<T>(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return default(T);
            }

            var json = File.ReadAllText(filePath);
            return JsonSerializer.Deserialize<T>(json, _options);
        }
    }
}