using System;
using System.IO;

namespace VBASinc.Sync
{
    public class SyncConfiguration
    {
        public SyncConfiguration()
        {
            RootFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "VBASinc");
        }

        public string RootFolderPath { get; set; }
        public bool SyncEnabled { get; set; } = true;
        public bool EnableExternalWatchers { get; set; } = true;
        public bool EnableInternalPolling { get; set; } = true;
        public int PollingIntervalSeconds { get; set; } = 14400;
        public bool SuppressSelfTriggeredEvents { get; set; } = true;
        public int SelfChangeSuppressionWindowMs { get; set; } = 5000;
        public int ConflictDiscoveryWindow { get; set; } = 2000;
        public bool AutoResolveConflicts { get; set; } = false;
        public bool PromptOnConflict { get; set; } = true;
        public string BackupFolderName { get; set; } = "Backups";
        public string[] IncludedExtensions { get; set; } = { ".bas", ".cls", ".frm", ".frx" };
        public string[] ExcludedModules { get; set; } = Array.Empty<string>();

        public void ApplyFrom(SyncConfiguration source)
        {
            if (source == null)
            {
                return;
            }

            RootFolderPath = source.RootFolderPath;
            SyncEnabled = source.SyncEnabled;
            EnableExternalWatchers = source.EnableExternalWatchers;
            EnableInternalPolling = source.EnableInternalPolling;
            PollingIntervalSeconds = source.PollingIntervalSeconds;
            SuppressSelfTriggeredEvents = source.SuppressSelfTriggeredEvents;
            SelfChangeSuppressionWindowMs = source.SelfChangeSuppressionWindowMs;
            ConflictDiscoveryWindow = source.ConflictDiscoveryWindow;
            AutoResolveConflicts = source.AutoResolveConflicts;
            PromptOnConflict = source.PromptOnConflict;
            BackupFolderName = source.BackupFolderName;
            IncludedExtensions = source.IncludedExtensions;
            ExcludedModules = source.ExcludedModules;
        }
    }
}
