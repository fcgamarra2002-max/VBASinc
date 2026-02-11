using System;
using System.Collections.Generic;
using System.IO;
using VBIDE = Microsoft.Vbe.Interop;

namespace VBASinc.Sync
{
    /// <summary>
    /// Contexto de sincronización para un proyecto VBA individual.
    /// Encapsula toda la información necesaria para mantener la sincronización de un libro.
    /// </summary>
    public class ProjectSyncContext : IDisposable
    {
        public string ProjectName { get; set; } = string.Empty;
        public string WorkbookPath { get; set; } = string.Empty;
        public string ExternalPath { get; set; } = string.Empty;
        public VBIDE.VBProject? VbaProject { get; set; }
        
        // Estado de sincronización
        public bool IsSyncEnabled { get; set; } = false;
        public DateTime LastSyncTime { get; set; } = DateTime.MinValue;
        
        // Hashes para detectar cambios
        public Dictionary<string, string> VbaHashes { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> FileHashes { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        
        // Estadísticas
        public int ExportCount { get; set; } = 0;
        public int ImportCount { get; set; } = 0;
        
        // Watcher individual
        public FileSystemWatcher? Watcher { get; set; }

        public void ClearHashes()
        {
            VbaHashes.Clear();
            FileHashes.Clear();
        }

        public void Dispose()
        {
            if (Watcher != null)
            {
                Watcher.EnableRaisingEvents = false;
                Watcher.Dispose();
                Watcher = null;
            }
        }

        public override string ToString()
        {
            string status = IsSyncEnabled ? "🟢" : "🔴";
            return $"{status} {ProjectName}";
        }
    }
}
