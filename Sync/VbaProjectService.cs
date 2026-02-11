using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using VBIDE = Microsoft.Vbe.Interop;

namespace VBASinc.Sync
{
    public class VbaProjectService
    {
        private readonly VBIDE.VBE _vbe;
        private readonly SyncConfiguration _configuration;

        public VbaProjectService(VBIDE.VBE vbe, SyncConfiguration configuration)
        {
            _vbe = vbe ?? throw new ArgumentNullException(nameof(vbe));
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
        }

        internal VBIDE.VBProject GetActiveProject()
        {
            if (TryGetActiveProject(out var project))
            {
                return project;
            }

            throw new InvalidOperationException("No se encontró ningún proyecto VBA cargado.");
        }

        internal bool TryGetActiveProject(out VBIDE.VBProject project)
        {
            project = null!;

            try
            {
                var active = _vbe.ActiveVBProject;
                if (active != null)
                {
                    project = active;
                    return true;
                }
            }
            catch
            {
                // ignorar y seguir intentos
            }

            try
            {
                if (_vbe.VBProjects != null && _vbe.VBProjects.Count > 0)
                {
                    project = _vbe.VBProjects.Item(1);
                    return project != null;
                }
            }
            catch
            {
                // ignorar
            }

            project = null!;
            return false;
        }

        public IEnumerable<VBIDE.VBComponent> GetModules()
        {
            if (TryGetActiveProject(out var project))
            {
                return project.VBComponents.Cast<VBIDE.VBComponent>();
            }

            return Enumerable.Empty<VBIDE.VBComponent>();
        }

        public VBIDE.VBComponent? FindModuleByName(string moduleName)
        {
            if (string.IsNullOrWhiteSpace(moduleName))
            {
                return null;
            }

            try
            {
                var project = GetActiveProject();
                return project.VBComponents.Item(moduleName);
            }
            catch
            {
                return null;
            }
        }

        public string ReadModuleCode(VBIDE.VBComponent component)
        {
            if (component == null)
            {
                throw new ArgumentNullException(nameof(component));
            }

            VBIDE.CodeModule codeModule = component.CodeModule;
            if (codeModule.CountOfLines <= 0)
            {
                return string.Empty;
            }

            return codeModule.Lines[1, codeModule.CountOfLines];
        }

        public void ReplaceModuleCode(VBIDE.VBComponent component, string code)
        {
            if (component == null)
            {
                throw new ArgumentNullException(nameof(component));
            }

            if (!string.IsNullOrEmpty(code))
            {
                // Remove Attribute lines appearing in the body, as they are metadata and cause syntax errors when added via AddFromString
                var lines = code.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                var filtered = lines.Where(l => !l.TrimStart().StartsWith("Attribute VB_", StringComparison.OrdinalIgnoreCase));
                code = string.Join(Environment.NewLine, filtered);
            }

            VBIDE.CodeModule codeModule = component.CodeModule;
            if (codeModule.CountOfLines > 0)
            {
                codeModule.DeleteLines(1, codeModule.CountOfLines);
            }

            if (!string.IsNullOrEmpty(code))
            {
                codeModule.AddFromString(code);
            }
        }

        public VBIDE.VBComponent? EnsureModule(string moduleName, VBIDE.vbext_ComponentType type)
        {
            var existing = FindModuleByName(moduleName);
            if (existing != null)
            {
                return existing;
            }

            if (!TryGetActiveProject(out var project))
            {
                return null;
            }

            var component = project.VBComponents.Add(type);

            if (!string.IsNullOrWhiteSpace(moduleName))
            {
                try
                {
                    component.Name = moduleName;
                }
                catch
                {
                    // Ignorar nombres inválidos; el host impondrá uno válido
                }
            }

            return component;
        }

        public void ExportModule(VBIDE.VBComponent component, string destinationPath)
        {
            if (component == null)
            {
                throw new ArgumentNullException(nameof(component));
            }

            string directory = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            component.Export(destinationPath);
        }

        public VBIDE.VBComponent? ImportFormModule(string frmPath)
        {
            if (string.IsNullOrWhiteSpace(frmPath))
            {
                throw new ArgumentNullException(nameof(frmPath));
            }

            if (!File.Exists(frmPath))
            {
                throw new FileNotFoundException("Form file not found", frmPath);
            }

            string moduleName = Path.GetFileNameWithoutExtension(frmPath);
            if (!TryGetActiveProject(out var project))
            {
                return null;
            }

            var existing = FindModuleByName(moduleName);
            if (existing != null)
            {
                project.VBComponents.Remove(existing);
            }

            return project.VBComponents.Import(frmPath);
        }

        public void RemoveModule(string moduleName)
        {
            if (string.IsNullOrWhiteSpace(moduleName))
            {
                return;
            }

            var existing = FindModuleByName(moduleName);
            if (existing != null)
            {
                if (TryGetActiveProject(out var project))
                {
                    project.VBComponents.Remove(existing);
                }
            }
        }

        internal ModuleSnapshot CaptureSnapshot(VBIDE.VBComponent component)
        {
            if (component == null)
            {
                throw new ArgumentNullException(nameof(component));
            }

            if (component.Type == VBIDE.vbext_ComponentType.vbext_ct_MSForm)
            {
                return CaptureFormSnapshot(component);
            }

            string content = ReadModuleCode(component);
            return new ModuleSnapshot(component.Name, component.Type, content, Array.Empty<byte>());
        }

        private ModuleSnapshot CaptureFormSnapshot(VBIDE.VBComponent component)
        {
            string tempRoot = Path.Combine(Path.GetTempPath(), "VBASinc", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempRoot);

            string frmPath = Path.Combine(tempRoot, component.Name + ".frm");

            try
            {
                component.Export(frmPath);
                string frmContent = File.ReadAllText(frmPath);
                string frxPath = Path.ChangeExtension(frmPath, ".frx");
                byte[] frxBytes = File.Exists(frxPath) ? File.ReadAllBytes(frxPath) : Array.Empty<byte>();
                return new ModuleSnapshot(component.Name, component.Type, frmContent, frxBytes);
            }
            finally
            {
                try
                {
                    if (Directory.Exists(tempRoot))
                    {
                        Directory.Delete(tempRoot, true);
                    }
                }
                catch
                {
                    // Ignorar errores de limpieza
                }
            }
        }

        internal readonly struct ModuleSnapshot
        {
            public ModuleSnapshot(string name, VBIDE.vbext_ComponentType componentType, string primaryContent, byte[]? binaryContent)
            {
                Name = name;
                ComponentType = componentType;
                PrimaryContent = primaryContent ?? string.Empty;
                BinaryContent = binaryContent ?? Array.Empty<byte>();
            }

            public string Name { get; }
            public VBIDE.vbext_ComponentType ComponentType { get; }
            public string PrimaryContent { get; }
            public byte[] BinaryContent { get; }
        }
    }
}
