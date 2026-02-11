using System;
using System.Runtime.InteropServices;
using Extensibility;

namespace VBASinc
{
    [ComVisible(true)]
    [Guid("AF5E8B64-56AC-4E23-9E91-EB8A8958A3F4")]
    [ProgId("VBASinc.Connect")]
    public class Connect : IDTExtensibility2
    {
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            // Minimal - do nothing
            try
            {
                System.IO.File.WriteAllText(@"C:\SincVBA\LOADED_OK.txt", "Add-in loaded at " + DateTime.Now.ToString());
            }
            catch { }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }
    }
}
