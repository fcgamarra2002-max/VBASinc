using System;
using System.Runtime.InteropServices;
using VBIDE = Microsoft.Vbe.Interop;

namespace VBASinc
{
    /// <summary>
    /// VBASinc - Self-Contained Synchronization System.
    /// Single entry point: Launch()
    /// Everything else is encapsulated.
    /// </summary>
    [ComVisible(true)]
    [Guid("A0B1C2D3-E4F5-6789-ABCD-EF0123456789")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("VBASinc.System")]
    public sealed class VBASincSystem : IVBASincSystem
    {
        private const string DEFAULT_PATH = @"C:\src_vba";

        /// <summary>
        /// Launches the complete synchronization system.
        /// This is the ONLY method exposed to VBA.
        /// </summary>
        /// <param name="vbaProject">VBProject COM object from ThisWorkbook.VBProject</param>
        public void Launch(object vbaProject)
        {
            Launch(vbaProject, DEFAULT_PATH);
        }

        /// <summary>
        /// Launches with custom external path.
        /// </summary>
        public void Launch(object vbaProject, string externalPath)
        {
            var project = vbaProject as VBIDE.VBProject;
            if (project == null)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Error: Objeto VBProject inválido.\n\nUso correcto:\nCreateObject(\"VBASinc.System\").Launch ThisWorkbook.VBProject",
                    "VBASinc",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return;
            }

            // Self-contained execution with guaranteed cleanup
            using (var form = new UI.SyncControlForm(project, externalPath ?? DEFAULT_PATH))
            {
                form.ShowDialog();
            }

            // Force cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    /// <summary>
    /// Minimal COM interface - only Launch() exposed.
    /// </summary>
    [ComVisible(true)]
    [Guid("B1C2D3E4-F567-890A-BCDE-F01234567890")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IVBASincSystem
    {
        void Launch(object vbaProject);
        void Launch(object vbaProject, string externalPath);
    }
}
