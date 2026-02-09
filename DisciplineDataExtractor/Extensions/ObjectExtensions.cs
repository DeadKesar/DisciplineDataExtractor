using System;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

namespace DisciplineDataExtractor.Extensions
{
    public static class ObjectExtensions
    {
        public static bool TryJsonSerialize(this object obj)
        {
            try
            {
                File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/log.json",
                    JsonSerializer.Serialize(obj, new JsonSerializerOptions
                    {
                        Encoder = JavaScriptEncoder.Create(UnicodeRanges.Cyrillic, UnicodeRanges.BasicLatin),
                        WriteIndented = true
                    }));
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
