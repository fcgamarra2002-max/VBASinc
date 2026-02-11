using System;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Web.Script.Serialization;

namespace VBASinc.Sync
{
    public class SyncSettingsStore
    {
        private const string ConfigFileName = "VBASincSettings.json";
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer();

        public SyncConfiguration Load()
        {
            string path = GetConfigFilePath();
            if (!File.Exists(path))
            {
                var configuration = new SyncConfiguration();
                ApplyDefaults(configuration);
                return configuration;
            }

            try
            {
                string json = File.ReadAllText(path);
                var config = _serializer.Deserialize<SyncConfiguration>(json) ?? new SyncConfiguration();
                ApplyDefaults(config);
                return config;
            }
            catch
            {
                var fallback = new SyncConfiguration();
                ApplyDefaults(fallback);
                return fallback;
            }
        }

        public void Save(SyncConfiguration configuration)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException(nameof(configuration));
            }

            var existing = Load();
            existing.ApplyFrom(configuration);
            string json = _serializer.Serialize(existing);
            string path = GetConfigFilePath();
            string directory = Path.GetDirectoryName(path)!;
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(path, json);
        }

        private static string GetConfigFilePath()
        {
            string roaming = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = Path.Combine(roaming, "VBASinc");
            Directory.CreateDirectory(folder);
            return Path.Combine(folder, ConfigFileName);
        }

        private void ApplyDefaults(SyncConfiguration configuration)
        {
            var settings = ConfigurationManager.AppSettings;
            if (settings == null)
            {
                return;
            }

            configuration.RootFolderPath = ResolvePath(settings["DefaultRootFolder"], configuration.RootFolderPath);
            configuration.PollingIntervalSeconds = GetInt(settings["DefaultPollingIntervalSeconds"], configuration.PollingIntervalSeconds);
            configuration.EnableExternalWatchers = GetBool(settings["DefaultEnableExternalWatchers"], configuration.EnableExternalWatchers);
            configuration.EnableInternalPolling = GetBool(settings["DefaultEnableInternalPolling"], configuration.EnableInternalPolling);
            configuration.SuppressSelfTriggeredEvents = GetBool(settings["DefaultSuppressSelfTriggeredEvents"], configuration.SuppressSelfTriggeredEvents);
            configuration.SelfChangeSuppressionWindowMs = GetInt(settings["DefaultSelfChangeSuppressionWindowMs"], configuration.SelfChangeSuppressionWindowMs);
            configuration.ConflictDiscoveryWindow = GetInt(settings["DefaultConflictDiscoveryWindow"], configuration.ConflictDiscoveryWindow);
            configuration.AutoResolveConflicts = GetBool(settings["DefaultAutoResolveConflicts"], configuration.AutoResolveConflicts);
            configuration.PromptOnConflict = GetBool(settings["DefaultPromptOnConflict"], configuration.PromptOnConflict);
            configuration.BackupFolderName = settings["DefaultBackupFolderName"] ?? configuration.BackupFolderName;
        }

        private static string ResolvePath(string? value, string fallback)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return fallback;
            }

            string expanded = Environment.ExpandEnvironmentVariables(value!.Trim());
            return string.IsNullOrWhiteSpace(expanded) ? fallback : expanded;
        }

        private static bool GetBool(string? value, bool fallback)
        {
            return bool.TryParse(value, out bool result) ? result : fallback;
        }

        private static int GetInt(string? value, int fallback)
        {
            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result))
            {
                return result;
            }

            return fallback;
        }
    }
}
