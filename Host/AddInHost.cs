using System;
using System.IO;
using Extensibility;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Collections.Concurrent;
using System.Collections.Generic;
using VBIDE = Microsoft.Vbe.Interop;
using System.Text;
using VBASinc.Diagnostics;
using VBASinc.UI;
using VBASinc.Sync;
using System.Linq;

namespace VBASinc.Host
{
    public class AddInHost : IDisposable
    {
        private VBIDE.VBE _vbe;
        private VBIDE.VBProject? _activeProject;
        private SyncControlForm? _syncForm;
        // private Office.CommandBarButton? _menuButton; // Removed dependency
        
        // Auto-Sync Components
        private FileSystemWatcher? _watcher;
        private System.Windows.Forms.Timer? _pollTimer;
        private readonly ConcurrentQueue<string> _pendingImports = new ConcurrentQueue<string>();
        private readonly Dictionary<string, string> _vbaHashes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _fileHashes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        
        // Configuration
        private string _externalPath = @"c:\SincVBA";
        private bool _syncEnabled = false;
        private bool _isInternalOperation = false;
        
        // Debounce
        private readonly Dictionary<string, DateTime> _lastEventTime = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
        private const int DEBOUNCE_MS = 0; // Instantaneous sync
        private readonly object _syncLock = new object();
        
        // Batch processing for large projects (1000+ modules)
        private int _exportBatchIndex = 0;
        private const int BATCH_SIZE = 10; // Process 10 modules per tick
        
        // Multi-Project Support
        private readonly Dictionary<string, ProjectSyncContext> _projectContexts = 
            new Dictionary<string, ProjectSyncContext>(StringComparer.OrdinalIgnoreCase);
        
        // Settings persistence
        private static readonly string ConfigPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "VBASinc", "config.txt");

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        public VBIDE.VBE VBE => _vbe;

        public AddInHost(VBIDE.VBE vbe)
        {
            _vbe = vbe;
        }

        public void Initialize()
        {
            AddInLogger.Log("AddInHost Initialized (Clean Version)");
        }

        public void Start()
        {
            try
            {
                // Removed early return to allow Menu setup
                try { _activeProject = _vbe.ActiveVBProject; } catch { }
                
                if (_activeProject == null)
                    WriteStatus("WARNING: No Active Project at Start. Waiting for project...");
                else
                    WriteStatus($"Startup Project: {_activeProject.Name}");
                
                // Load saved path or use default
                LoadSavedPath();
                if (!Directory.Exists(_externalPath)) Directory.CreateDirectory(_externalPath);

                SetupMenuProtected();
                // ShowUI(); // COMENTADO: No abrir automáticamente al entrar en modo desarrollador
            }
            catch (Exception ex)
            {
                AddInLogger.Log("Error starting AddInHost", ex);
                WriteStatus("CRITICAL START FAILURE: " + ex.ToString());
            }
        }

        private VBIDE.CommandBarEvents? _menuEvents;

        private void SetupMenuProtected()
        {
            try
            {
                 WriteStatus("SetupMenuProtected: Starting...");
                 SetupMenuReflectionActual();
                 WriteStatus("SetupMenuProtected: Success.");
            }
            catch (Exception ex)
            {
                 // Do not log to AddInLogger here if it might fail, rely on WriteStatus
                 WriteStatus("MENU SETUP FAILED: " + ex.ToString());
            }
        }

        private void SetupMenuReflectionActual()
        {
            try
            {
                WriteStatus("SetupMenuReflectionActual: Starting...");

                // Use Reflection to avoid Office dependency on CommandBars property type
                object commandBars = _vbe.GetType().InvokeMember("CommandBars", System.Reflection.BindingFlags.GetProperty, null, _vbe, null);
                object? menuBar = null;

                // 1. Try "Menu Bar"
                try {
                    WriteStatus("Trying to get 'Menu Bar'...");
                    menuBar = GetProperty(commandBars, "Item", "Menu Bar");
                } catch {
                    WriteStatus("'Menu Bar' not found via Item property.");
                }

                // 2. Try ActiveMenuBar
                if (menuBar == null)
                {
                    try {
                        WriteStatus("Trying to get ActiveMenuBar...");
                        menuBar = GetProperty(commandBars, "ActiveMenuBar");
                    } catch {
                        WriteStatus("ActiveMenuBar not found.");
                    }
                }

                if (menuBar == null)
                {
                    // 3. Search by Type (MsoBarType.msoBarTypeMenuBar = 1)
                     // commandBars is IEnumerable
                    System.Collections.IEnumerable? bars = commandBars as System.Collections.IEnumerable;
                    if (bars != null)
                    {
                        WriteStatus("Searching menu bar by type...");
                        foreach (object bar in bars)
                        {
                            try {
                                object typeObj = GetProperty(bar, "Type");
                                object visObj = GetProperty(bar, "Visible");
                                if (Convert.ToInt32(typeObj) == 1 && Convert.ToBoolean(visObj))
                                {
                                    menuBar = bar;
                                    WriteStatus("Menu bar found by type!");
                                    break;
                                }
                            } catch { }
                        }
                    }
                }

                if (menuBar == null)
                {
                    WriteStatus("MENU ERROR: Menu Bar not found via Reflection.");
                    return;
                }

                string menuBarName = "Unknown";
                try { menuBarName = GetProperty(menuBar, "Name")?.ToString() ?? "Unknown"; } catch {}
                WriteStatus($"Menu Bar found: {menuBarName}");

                // Cleanup "SincVBA" - try to remove any existing button first
                try {
                    WriteStatus("Attempting to remove existing SincVBA button...");
                    object controls = GetProperty(menuBar, "Controls");
                    object old = GetProperty(controls, "Item", "SincVBA");
                    if (old != null)
                    {
                        InvokeMethod(old, "Delete");
                        WriteStatus("Existing SincVBA button removed.");
                    }
                } catch {
                    WriteStatus("No existing SincVBA button found to remove.");
                }

                // Add Button directly to menu bar (not as a submenu)
                object barControls = GetProperty(menuBar, "Controls");
                // Add(Type=1, Id, Parameter, Before, Temporary=false) - Type=1 is a button
                object btn = InvokeMethod(barControls, "Add", 1, Type.Missing, Type.Missing, Type.Missing, false);

                // --- CONFIGURACIÓN DE VISIBILIDAD DE TEXTO CENTRADO ---
                // Set style = 2 (msoButtonCaption) for centered text-only button
                // This centers the text without an icon area
                SetProperty(btn, "Style", 2);
                
                // Set Caption (must be after Style for proper rendering)
                SetProperty(btn, "Caption", "SincVBA");
                
                // Unique identifier
                SetProperty(btn, "Tag", "VBASincButton");
                
                // Ensure visibility and enable
                SetProperty(btn, "Visible", true);
                SetProperty(btn, "Enabled", true);

                // Tooltip
                try { SetProperty(btn, "TooltipText", "Abrir Sincronizador VBA"); } catch {}
                
                // Separate from standard menus
                try { SetProperty(btn, "BeginGroup", true); } catch { }

                WriteStatus("Button added to menu bar: SincVBA (Style 2 = Caption only, centered)");

                // Hook Event using VBIDE Events (No Office Reference Needed)
                try {
                    _menuEvents = _vbe.Events.get_CommandBarEvents(btn);
                    _menuEvents.Click += OnMenuClick;
                    WriteStatus("Event hooked successfully via VBIDE.");
                }
                catch (Exception castEx)
                {
                   WriteStatus("MENU EVENT HOOK FAILED: " + castEx.Message);
                }

                WriteStatus("SetupMenuReflectionActual: Completed successfully.");

            }
            catch (Exception ex)
            {
                WriteStatus("Reflection Menu Error: " + ex.Message);
                throw new Exception("Reflection Menu Error: " + ex.Message, ex);
            }
        }

        private void OnMenuClick(object CommandBarControl, ref bool handled, ref bool CancelDefault)
        {
            ShowUI();
            handled = true;
            CancelDefault = true;
        }

        // Reflection Helpers
        private object GetProperty(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, target, args);
        }

        private void SetProperty(object target, string name, object value)
        {
            target.GetType().InvokeMember(name, System.Reflection.BindingFlags.SetProperty, null, target, new object[] { value });
        }

        private object InvokeMethod(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(name, System.Reflection.BindingFlags.InvokeMethod, null, target, args);
        }

        private void ShowUI()
        {
            try
            {
                // Always try to refresh active project to handle new workbooks or switching
                try 
                { 
                    var currentActive = _vbe.ActiveVBProject; 
                    if (currentActive != null) _activeProject = currentActive;
                } 
                catch { }

                if (_syncForm == null || _syncForm.IsDisposed)
                {
                    // Create form even if project is still null (form will handle it or show error)
                    _syncForm = new SyncControlForm(_activeProject, _externalPath, this);
                }
                else if (_syncForm.VbaProject == null && _activeProject != null)
                {
                    // Update project in existing form if it was null before
                    _syncForm.VbaProject = _activeProject;
                }

                if (_syncForm != null)
                {
                    if (_syncForm.WindowState == FormWindowState.Minimized)
                        _syncForm.WindowState = FormWindowState.Normal;

                    _syncForm.Show();
                    _syncForm.BringToFront();
                    _syncForm.Focus();
                }
            }
            catch (Exception ex)
            {
                WriteStatus("ShowUI Error: " + ex.Message);
            }
        }

        public bool IsSyncEnabled => _syncEnabled;
        public string ExternalPath => _externalPath;

        public void SetExternalPath(string newPath)
        {
            if (string.IsNullOrEmpty(newPath)) return;
            if (!Directory.Exists(newPath)) Directory.CreateDirectory(newPath);
            
            bool wasEnabled = _syncEnabled;
            
            // Stop current sync if running
            if (wasEnabled) StopBackgroundSync();
            
            // Update path
            _externalPath = newPath;
            _vbaHashes.Clear();
            _fileHashes.Clear();
            
            WriteStatus($"PATH CHANGED: {newPath}");
            
            // Save path for next session
            SavePath(newPath);
            
            // Restart sync if it was running
            if (wasEnabled) StartBackgroundSync();
        }

        private void LoadSavedPath()
        {
            try
            {
                AddInLogger.Log($"LoadSavedPath: Looking for config at {ConfigPath}");
                
                if (File.Exists(ConfigPath))
                {
                    string savedPath = File.ReadAllText(ConfigPath).Trim();
                    AddInLogger.Log($"LoadSavedPath: Config found, content = '{savedPath}'");
                    
                    if (!string.IsNullOrEmpty(savedPath))
                    {
                        // Create directory if it doesn't exist
                        if (!Directory.Exists(savedPath))
                        {
                            try { Directory.CreateDirectory(savedPath); } catch { }
                        }
                        
                        _externalPath = savedPath;
                        AddInLogger.Log($"LoadSavedPath: Path SET to '{_externalPath}'");
                    }
                }
                else
                {
                    AddInLogger.Log($"LoadSavedPath: Config file NOT found, using default '{_externalPath}'");
                }
            }
            catch (Exception ex)
            {
                AddInLogger.Log($"LoadSavedPath ERROR: {ex.Message}");
            }
        }

        private void SavePath(string path)
        {
            try
            {
                string? dir = Path.GetDirectoryName(ConfigPath);
                if (dir != null && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
                
                File.WriteAllText(ConfigPath, path);
            }
            catch { }
        }

        public void StartBackgroundSync()
        {
            if (_syncEnabled) return;

            try
            {
                // Init Watcher
                _watcher = new FileSystemWatcher(_externalPath)
                {
                    Filter = "*.*", 
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.Size | NotifyFilters.Attributes,
                    IncludeSubdirectories = true,
                    InternalBufferSize = 4194304 // 4MB max for massive files
                };

                _watcher.Changed += OnExternalChange;
                _watcher.Created += OnExternalChange;
                _watcher.Renamed += OnExternalRename;
                _watcher.Error += OnWatcherError;
                _watcher.EnableRaisingEvents = true;

                // Init Timer (2000ms - sync every 2 seconds)
                _pollTimer = new System.Windows.Forms.Timer { Interval = 200 };
                _pollTimer.Tick += OnPollTick;
                _pollTimer.Start();

                _syncEnabled = true;
                _vbaHashes.Clear();
                _fileHashes.Clear();
                
                // Initial Sync (Export current state to be safe, or just start listening)
                // For safety, we just start listening.
                WriteStatus("AUTO-SYNC STARTED");
            }
            catch (Exception ex)
            {
                AddInLogger.Log("Failed to start sync", ex);
                WriteStatus($"ERROR STARTING: {ex.Message}");
            }
        }

        public void StopBackgroundSync()
        {
            _syncEnabled = false;
            if (_watcher != null)
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Dispose();
                _watcher = null;
            }
            if (_pollTimer != null)
            {
                _pollTimer.Stop();
                _pollTimer.Dispose();
                _pollTimer = null;
            }
            WriteStatus("AUTO-SYNC STOPPED");
        }

        #region Multi-Project Sync

        /// <summary>
        /// Inicia sincronización para un proyecto específico con su propia carpeta.
        /// </summary>
        public bool StartSyncForProject(VBIDE.VBProject project, string externalPath)
        {
            if (project == null) return false;
            
            string projectKey = GetProjectKey(project);
            
            // Si ya existe, detener primero
            if (_projectContexts.ContainsKey(projectKey))
                StopSyncForProject(project);
            
            try
            {
                var context = new ProjectSyncContext
                {
                    ProjectName = project.Name,
                    VbaProject = project,
                    ExternalPath = externalPath,
                    IsSyncEnabled = true,
                    LastSyncTime = DateTime.Now
                };
                
                // Crear carpeta si no existe
                if (!Directory.Exists(externalPath))
                    Directory.CreateDirectory(externalPath);
                
                // Crear watcher para este proyecto
                context.Watcher = new FileSystemWatcher(externalPath)
                {
                    Filter = "*.*",
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName,
                    IncludeSubdirectories = true,
                    InternalBufferSize = 4194304
                };
                
                context.Watcher.Changed += (s, e) => OnProjectExternalChange(projectKey, e.FullPath);
                context.Watcher.Created += (s, e) => OnProjectExternalChange(projectKey, e.FullPath);
                context.Watcher.EnableRaisingEvents = true;
                
                _projectContexts[projectKey] = context;
                
                WriteStatus($"SYNC STARTED: {project.Name}");
                return true;
            }
            catch (Exception ex)
            {
                AddInLogger.Log($"Failed to start sync for {project.Name}", ex);
                return false;
            }
        }

        /// <summary>
        /// Detiene sincronización para un proyecto específico.
        /// </summary>
        public void StopSyncForProject(VBIDE.VBProject project)
        {
            if (project == null) return;
            
            string projectKey = GetProjectKey(project);
            
            if (_projectContexts.TryGetValue(projectKey, out var context))
            {
                context.Dispose();
                _projectContexts.Remove(projectKey);
                WriteStatus($"SYNC STOPPED: {project.Name}");
            }
        }

        /// <summary>
        /// Verifica si un proyecto tiene sincronización activa.
        /// </summary>
        public bool IsProjectSyncing(VBIDE.VBProject project)
        {
            if (project == null) return false;
            string projectKey = GetProjectKey(project);
            return _projectContexts.ContainsKey(projectKey) && 
                   _projectContexts[projectKey].IsSyncEnabled;
        }

        /// <summary>
        /// Obtiene la ruta de sincronización de un proyecto.
        /// </summary>
        public string? GetProjectSyncPath(VBIDE.VBProject project)
        {
            if (project == null) return null;
            string projectKey = GetProjectKey(project);
            return _projectContexts.TryGetValue(projectKey, out var ctx) ? ctx.ExternalPath : null;
        }

        /// <summary>
        /// Obtiene lista de todos los proyectos con sincronización activa.
        /// </summary>
        public List<string> GetActiveSyncProjects()
        {
            return _projectContexts.Where(p => p.Value.IsSyncEnabled)
                                   .Select(p => p.Value.ProjectName)
                                   .ToList();
        }

        private string GetProjectKey(VBIDE.VBProject project)
        {
            try
            {
                return $"{project.Name}_{project.FileName?.GetHashCode() ?? 0}";
            }
            catch
            {
                return project.Name;
            }
        }

        private void OnProjectExternalChange(string projectKey, string fullPath)
        {
            if (_isInternalOperation) return;
            if (!_projectContexts.TryGetValue(projectKey, out var context)) return;
            if (context.VbaProject == null) return;
            
            string filename = Path.GetFileName(fullPath);
            if (filename.StartsWith("~") || !IsSupportedExtension(Path.GetExtension(filename)))
                return;
            
            // Encolar para importación en el contexto de este proyecto
            _pendingImports.Enqueue($"{projectKey}|{fullPath}");
        }

        #endregion

        #region Event Handlers

        private void OnExternalChange(object sender, FileSystemEventArgs e) => EnqueueChange(e.FullPath);
        
        private void OnExternalRename(object sender, RenamedEventArgs e) => EnqueueChange(e.FullPath);

        private void EnqueueChange(string fullPath)
        {
            if (_isInternalOperation) return;
            
            string filename = Path.GetFileName(fullPath);
            // Ignore temporary files, status files, and logs
            if (filename.StartsWith("~") || 
                filename.Equals("vbasinc_status.txt", StringComparison.OrdinalIgnoreCase) ||
                filename.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase) ||
                filename.EndsWith(".log", StringComparison.OrdinalIgnoreCase)) 
                return;

            if (!IsSupportedExtension(Path.GetExtension(filename))) return;

            // Debounce
            lock (_syncLock)
            {
                DateTime now = DateTime.UtcNow;
                if (_lastEventTime.TryGetValue(fullPath, out DateTime last) && (now - last).TotalMilliseconds < DEBOUNCE_MS)
                    return;
                _lastEventTime[fullPath] = now;
            }

            _pendingImports.Enqueue(fullPath);
        }

        private void OnWatcherError(object sender, ErrorEventArgs e)
        {
            // Auto-restart watcher logic could go here, for now just log
            AddInLogger.Log("Watcher Error", e.GetException());
            // Restarting is risky if done recursively. Better to just let it fail or restart cleanly.
            // For now, let's try to restart once on next tick if needed.
        }

        // Cooldown to prevent export right after import (avoids conflict)
        private DateTime _lastImportTime = DateTime.MinValue;
        private const int EXPORT_COOLDOWN_MS = 100;

        private void OnPollTick(object? sender, EventArgs e)
        {
            try
            {
                // 0. Auto-Connect Project if missing
                if (_activeProject == null)
                {
                    try { 
                        _activeProject = _vbe.ActiveVBProject; 
                        if (_activeProject != null) 
                        {
                            WriteStatus($"Project Connected: {_activeProject.Name}");
                        }
                    } catch { }
                }

                // 1. Process Imports (HIGH PRIORITY - always first)
                bool didImport = false;
                HashSet<string> processed = new HashSet<string>();
                while (_pendingImports.TryDequeue(out string? path))
                {
                    if (path != null && processed.Add(path))
                    {
                        WriteStatus($"IMPORTING: {Path.GetFileName(path)}");
                        ExecuteImport(path);
                        didImport = true;
                    }
                }

                // If we just imported, set cooldown and skip export this tick
                if (didImport)
                {
                    _lastImportTime = DateTime.UtcNow;
                    return;
                }

                // 2. Process Exports (ALWAYS - no VBE restriction)
                // Skip if still in cooldown period after import
                if ((DateTime.UtcNow - _lastImportTime).TotalMilliseconds < EXPORT_COOLDOWN_MS)
                    return;

                // Export changes from VBA to disk (works even with VBE open)
                ExecuteExport();
            }
            catch (Exception ex)
            {
                AddInLogger.Log("Poll Error", ex);
            }
        }

        #endregion

        #region Core Logic

        private void ExecuteImport(string filePath)
        {
            if (!File.Exists(filePath)) return;
            if (_activeProject == null) return;

            string name = Path.GetFileNameWithoutExtension(filePath);
            
            try
            {
                _isInternalOperation = true; // Prevent feedback loop starts here

                // Read Content (Robust)
                string? content = ReadFileRobust(filePath);
                if (content == null) return; // Read failed

                // Compute Hash - compare against PREVIOUS FILE hash, not VBA hash
                string newHash = ComputeHash(content);
                if (_fileHashes.TryGetValue(name, out string? oldFileHash) && oldFileHash == newHash) 
                    return; // External file unchanged

                // Update in VBA
                VBIDE.VBComponent? comp = null;
                try { comp = _activeProject.VBComponents.Item(name); } catch { }

                if (comp == null)
                {
                    // Create new
                    // Determine type from extension is tricky without parsing, 
                    // for now assume standard module if not found, or try to infer.
                    // Ideally we should map extension to type.
                    var type = ExtensionToType(Path.GetExtension(filePath));
                    comp = _activeProject.VBComponents.Add(type);
                    comp.Name = name;
                }

                if (comp.CodeModule.CountOfLines > 0)
                    comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines);
                
                comp.CodeModule.AddFromString(content);

                // Update both hashes (file was imported, now VBA matches file)
                _fileHashes[name] = newHash;
                _vbaHashes[name] = newHash;
                WriteStatus($"IMPORTED: {name}");

                // Force UI Refresh (Visual) - Multiple techniques for immediate visibility
                try { 
                    var codePane = comp.CodeModule.CodePane;
                    if (codePane != null)
                    {
                        // Force scroll to top then back to trigger visual refresh
                        int topLine = 1;
                        try { topLine = codePane.TopLine; } catch { }
                        
                        codePane.TopLine = 1;
                        codePane.TopLine = topLine > 1 ? topLine : 1;
                        
                        // Set focus to force repaint
                        try { codePane.Window.SetFocus(); } catch { }
                        
                        // Alternative: Show the pane if hidden
                        try { codePane.Show(); } catch { }
                    }
                } catch { }

                // Notify Form if open
                if (_syncForm != null && !_syncForm.IsDisposed)
                    _syncForm.PublicRefreshUI(name, "IMPORTADO");

            }
            catch (Exception ex)
            {
                WriteStatus($"ERROR IMPORT: {name} - {ex.Message}");
            }
            finally
            {
                _isInternalOperation = false;
            }
        }

        private void ExecuteExport()
        {
            if (_activeProject == null) return;

            try
            {
                // PRIORITY 1: Export ACTIVE module first (the one being edited)
                try
                {
                    var activePane = _vbe.ActiveCodePane;
                    if (activePane != null)
                    {
                        var activeComp = activePane.CodeModule.Parent;
                        if (activeComp != null)
                        {
                            ExportSingleComponent(activeComp);
                        }
                    }
                }
                catch { } // No active pane, continue with batch

                // PRIORITY 2: Batch process other modules (rotating)
                var components = _activeProject.VBComponents.Cast<VBIDE.VBComponent>().ToArray();
                int total = components.Length;
                
                if (total == 0) return;
                if (_exportBatchIndex >= total) _exportBatchIndex = 0;
                
                int processed = 0;
                int startIndex = _exportBatchIndex;
                
                for (int i = 0; i < total && processed < BATCH_SIZE; i++)
                {
                    int idx = (startIndex + i) % total;
                    var comp = components[idx];
                    
                    // Skip if already exported as active module
                    ExportSingleComponent(comp);
                    processed++;
                }
                
                _exportBatchIndex = (_exportBatchIndex + BATCH_SIZE) % total;
            }
            catch { }
        }

        private void ExportSingleComponent(VBIDE.VBComponent comp)
        {
            try
            {
                string code = GetCleanCode(comp);
                string currentHash = ComputeHash(code);

                // Skip if unchanged
                if (_vbaHashes.TryGetValue(comp.Name, out string? oldVbaHash) && oldVbaHash == currentHash)
                    return;

                // Change detected - export!
                string ext = TypeToExtension(comp.Type);
                string subfolder = GetSubfolder(comp.Type);
                string targetDir = Path.Combine(_externalPath, subfolder);
                
                if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);
                
                string path = Path.Combine(targetDir, comp.Name + ext);

                try
                {
                    _isInternalOperation = true;
                    // Use UTF-8 with BOM for compatibility with modern editors (VS Code, etc.)
                    using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 65536))
                    using (var sw = new StreamWriter(fs, new UTF8Encoding(true), 65536)) // UTF-8 with BOM
                    {
                        sw.Write(code);
                    }
                    _vbaHashes[comp.Name] = currentHash;
                    _fileHashes[comp.Name] = currentHash;
                    
                    WriteStatus($"EXPORTED: {comp.Name}");
                    if (_syncForm != null && !_syncForm.IsDisposed)
                        _syncForm.PublicRefreshUI(comp.Name, "EXPORTADO");
                }
                finally { _isInternalOperation = false; }
            }
            catch { }
        }

        #endregion

        #region Helpers

        private string? ReadFileRobust(string path)
        {
            // Maximum retries for extremely large files (millions of lines)
            for (int i = 0; i < 50; i++)
            {
                try
                {
                    byte[] bytes = File.ReadAllBytes(path);
                    
                    // Detect UTF-8 BOM
                    if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                        return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
                    
                    // Detect UTF-8 without BOM (check for valid UTF-8 multi-byte sequences)
                    if (IsValidUtf8(bytes))
                        return Encoding.UTF8.GetString(bytes);
                    
                    // Fallback: Windows-1252 for legacy VBA files
                    return Encoding.GetEncoding(1252).GetString(bytes);
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(1); // Absolute minimum
                }
            }
            return null; // Failed after 50 attempts
        }

        private bool IsValidUtf8(byte[] bytes)
        {
            // Check for valid UTF-8 multi-byte sequences
            // Returns true if the file contains UTF-8 multi-byte chars (like Spanish accents in UTF-8)
            int i = 0;
            bool hasMultiByte = false;
            
            while (i < bytes.Length)
            {
                byte b = bytes[i];
                
                if (b <= 0x7F)
                {
                    // ASCII - valid in both encodings
                    i++;
                }
                else if (b >= 0xC2 && b <= 0xDF)
                {
                    // 2-byte UTF-8 sequence (covers á, é, í, ó, ú, ñ, etc.)
                    if (i + 1 >= bytes.Length || (bytes[i + 1] & 0xC0) != 0x80)
                        return false; // Invalid sequence, probably Windows-1252
                    hasMultiByte = true;
                    i += 2;
                }
                else if (b >= 0xE0 && b <= 0xEF)
                {
                    // 3-byte UTF-8 sequence
                    if (i + 2 >= bytes.Length || (bytes[i + 1] & 0xC0) != 0x80 || (bytes[i + 2] & 0xC0) != 0x80)
                        return false;
                    hasMultiByte = true;
                    i += 3;
                }
                else if (b >= 0xF0 && b <= 0xF4)
                {
                    // 4-byte UTF-8 sequence
                    if (i + 3 >= bytes.Length || (bytes[i + 1] & 0xC0) != 0x80 || (bytes[i + 2] & 0xC0) != 0x80 || (bytes[i + 3] & 0xC0) != 0x80)
                        return false;
                    hasMultiByte = true;
                    i += 4;
                }
                else
                {
                    // Invalid UTF-8 start byte (0x80-0xBF, 0xC0-0xC1, 0xF5-0xFF)
                    // This is likely Windows-1252 (single-byte extended chars)
                    return false;
                }
            }
            
            // Only return true if we found multi-byte sequences (actual UTF-8 chars)
            // Pure ASCII files should use Windows-1252 for VBA compatibility
            return hasMultiByte;
        }

        private string ComputeHash(string content)
        {
            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(content);
                return Convert.ToBase64String(sha.ComputeHash(bytes));
            }
        }

        private void WriteStatus(string msg)
        {
            try { File.AppendAllText(Path.Combine(_externalPath, "vbasinc_status.txt"), $"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}"); } catch { }
        }

        private bool IsVbeActive()
        {
            try
            {
                IntPtr fg = GetForegroundWindow();
                return fg == new IntPtr(_vbe.MainWindow.HWnd);
            }
            catch { return false; }
        }

        private string GetCleanCode(VBIDE.VBComponent comp)
        {
            var cm = comp.CodeModule;
            if (cm.CountOfLines == 0) return "";
            string code = cm.Lines[1, cm.CountOfLines];
            // Remove Attributes would go here if needed, but for simple sync we keep raw code usually 
            // OR use Regex to clean attributes if we are strictly code-focused. 
            // For now, return raw code to ensure consistency.
            return code;
        }

        private bool IsSupportedExtension(string ext)
        {
            ext = ext.ToLowerInvariant();
            return ext == ".bas" || ext == ".cls" || ext == ".frm";
        }
        
        private string TypeToExtension(VBIDE.vbext_ComponentType type)
        {
            switch (type)
            {
                case VBIDE.vbext_ComponentType.vbext_ct_ClassModule: return ".cls";
                case VBIDE.vbext_ComponentType.vbext_ct_MSForm: return ".frm";
                case VBIDE.vbext_ComponentType.vbext_ct_Document: return ".cls"; // ThisWorkbook, Sheets
                default: return ".bas";
            }
        }

        private string GetSubfolder(VBIDE.vbext_ComponentType type)
        {
            switch (type)
            {
                case VBIDE.vbext_ComponentType.vbext_ct_StdModule: return "Módulos";
                case VBIDE.vbext_ComponentType.vbext_ct_MSForm: return "Formularios";
                case VBIDE.vbext_ComponentType.vbext_ct_ClassModule: return "Módulos de clase";
                case VBIDE.vbext_ComponentType.vbext_ct_Document: return "Microsoft Excel Objetos";
                default: return "Otros";
            }
        }

        private VBIDE.vbext_ComponentType ExtensionToType(string ext)
        {
            switch (ext.ToLowerInvariant())
            {
                case ".cls": return VBIDE.vbext_ComponentType.vbext_ct_ClassModule;
                case ".frm": return VBIDE.vbext_ComponentType.vbext_ct_MSForm;
                default: return VBIDE.vbext_ComponentType.vbext_ct_StdModule;
            }
        }

        public void Dispose()
        {
            RemoveMenu();
            StopBackgroundSync();
            _syncForm?.Dispose();
        }

        // Legacy SetupMenu removed


        private void RemoveMenu()
        {
            try
            {
                if (_menuEvents != null)
                {
                    _menuEvents.Click -= OnMenuClick;
                    _menuEvents = null;
                }
                
                // Cleanup using Reflection fallback or simple try-catch generic logic
                // We cleaned up button logic in SetupMenu usually, but on Dispose we ideally clean up.
                // Since we don't hold the button ref anymore (it was local object), 
                // we can search and destroy or just leave it temporary (it is Temporary=true).
                // Temporary controls are auto-deleted by Office on restart, so manual delete is optional but good.
                
                // For simplicity and safety against Reference Errors, we rely on Temporary=true 
                // and the cleanup at start of SetupMenu.
            }
            catch { }
        }

        #endregion
    }
}
