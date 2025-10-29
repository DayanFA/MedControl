using System;
using System.IO;
using System.Text.Json;

namespace MedControl
{
    public class AppConfigModel
    {
        public string DbProvider { get; set; } = "sqlite"; // "sqlite" | "mysql"
        public string? MySqlConnectionString { get; set; }
    }

    public static class AppConfig
    {
        private static readonly string ConfigPath = Path.Combine(AppContext.BaseDirectory, "dbconfig.json");
        private static AppConfigModel _instance = new AppConfigModel();

        public static AppConfigModel Instance
        {
            get
            {
                if (_instance == null)
                {
                    Load();
                }
                return _instance;
            }
        }

        public static void Load()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    var json = File.ReadAllText(ConfigPath);
                    var model = JsonSerializer.Deserialize<AppConfigModel>(json, new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true
                    });
                    _instance = model ?? new AppConfigModel();
                }
                else
                {
                    _instance = new AppConfigModel();
                    Save();
                }
            }
            catch
            {
                _instance = new AppConfigModel();
            }
        }

        public static void Save()
        {
            try
            {
                var opts = new JsonSerializerOptions { WriteIndented = true };
                var json = JsonSerializer.Serialize(_instance, opts);
                File.WriteAllText(ConfigPath, json);
            }
            catch
            {
                // ignore
            }
        }
    }
}
