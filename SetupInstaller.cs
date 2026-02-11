using System;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Win32;

namespace VBASinc.Installer
{
    static class Program
    {
        // El destino ahora es AppData para no ensuciar el disco
        private static readonly string TargetFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
            "VBASinc"
        );

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (!IsAdministrator())
            {
                RestartAsAdmin();
                return;
            }

            try
            {
                RunInstallation();
                MessageBox.Show("VBASinc v1.0.4 se ha instalado y registrado correctamente en su carpeta de usuario (AppData).\n\nYa puede usar el complemento en Microsoft Office.", 
                    "Instalación Completada", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hubo un error durante la instalación:\n\n" + ex.Message, 
                    "Error de Instalación", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        static bool IsAdministrator()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        static void RestartAsAdmin()
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = Application.ExecutablePath,
                UseShellExecute = true,
                Verb = "runas"
            };

            try
            {
                Process.Start(startInfo);
            }
            catch (System.ComponentModel.Win32Exception)
            {
            }
        }

        static void RunInstallation()
        {
            // 1. Asegurar carpeta destino en AppData
            if (!Directory.Exists(TargetFolder))
            {
                Directory.CreateDirectory(TargetFolder);
            }

            // 2. Extraer recursos embebidos a AppData
            ExtractResource("VBASinc.dll", Path.Combine(TargetFolder, "VBASinc.dll"));
            ExtractResource("Extensibility.dll", Path.Combine(TargetFolder, "Extensibility.dll"));
            // El .bat no es necesario extraerlo si el instalador hace el trabajo, pero se guarda por si acaso
            ExtractResource("RegistrarComplemento.bat", Path.Combine(TargetFolder, "RegistrarComplemento.bat"));

            string dllPath = Path.Combine(TargetFolder, "VBASinc.dll");

            string regasm64 = @"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe";
            string regasm32 = @"C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe";

            // 3. Registro COM (Silencioso y desde AppData)
            Register(regasm64, dllPath);
            Register(regasm32, dllPath);

            // 4. Configuración de Add-in en el registro
            SetupAddInRegistry();

            // 5. Configuración de Seguridad (Apuntando a AppData)
            SetupSecurity();
        }

        static void ExtractResource(string resourceName, string outputPath)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    SaveStream(stream, outputPath);
                }
            }
        }

        static void SaveStream(Stream stream, string outputPath)
        {
            using (FileStream fileStream = new FileStream(outputPath, FileMode.Create))
            {
                stream.CopyTo(fileStream);
            }
        }

        static void SetupSecurity()
        {
            string[] apps = { "Excel", "Word", "PowerPoint" };

            foreach (string app in apps)
            {
                try
                {
                    // 1. App-specific Security
                    string securityKeyPath = string.Format(@"Software\Microsoft\Office\16.0\{0}\Security", app);
                    using (RegistryKey key = Registry.CurrentUser.CreateSubKey(securityKeyPath))
                    {
                        if (key != null)
                        {
                            key.SetValue("AccessVBOM", 1, RegistryValueKind.DWord);
                            key.SetValue("RequireAddInSig", 0, RegistryValueKind.DWord);
                        }
                    }

                    // 2. Trusted Location (Usando la nueva ruta de AppData)
                    string locationKeyPath = string.Format(@"Software\Microsoft\Office\16.0\{0}\Security\Trusted Locations\VBASinc", app);
                    using (RegistryKey key = Registry.CurrentUser.CreateSubKey(locationKeyPath))
                    {
                        if (key != null)
                        {
                            key.SetValue("Path", TargetFolder, RegistryValueKind.String);
                            key.SetValue("Description", "VBASinc AppData", RegistryValueKind.String);
                            key.SetValue("AllowSubfolders", 1, RegistryValueKind.DWord);
                        }
                    }
                }
                catch { }
            }
        }

        static void Register(string regasm, string dll)
        {
            if (File.Exists(regasm))
            {
                RunProcess(regasm, string.Format("\"{0}\" /codebase /tlb /silent", dll));
            }
        }

        static void RunProcess(string filename, string args)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = filename,
                Arguments = args,
                CreateNoWindow = true,
                UseShellExecute = false,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            try
            {
                using (var process = Process.Start(startInfo))
                {
                    if (process != null) process.WaitForExit();
                }
            }
            catch { }
        }

        static void SetupAddInRegistry()
        {
            string[] vbeVersions = { "6.0", "7.1" };
            string[] vbeTargets = { "Addins", "Addins64" };

            foreach (string v in vbeVersions)
            {
                foreach (string t in vbeTargets)
                {
                    try
                    {
                        string keyPath = string.Format(@"Software\Microsoft\VBA\VBE\{0}\{1}\VBASinc.Connect", v, t);
                        using (RegistryKey key = Registry.CurrentUser.CreateSubKey(keyPath))
                        {
                            if (key != null)
                            {
                                key.SetValue("FriendlyName", "VBASinc", RegistryValueKind.String);
                                key.SetValue("Description", "Sincronizador VBA (v1.0.4)", RegistryValueKind.String);
                                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                            }
                        }
                    }
                    catch { }
                }
            }
        }
    }
}
