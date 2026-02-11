using System;
using System.Runtime.InteropServices;
using Extensibility;
using VBASinc.Diagnostics;
using VBASinc.Host;
using VBASinc.UI;
using VBIDE = Microsoft.Vbe.Interop;

namespace VBASinc
{
    [ComVisible(true)]
    [Guid("AF5E8B64-56AC-4E23-9E91-EB8A8958A3F4")]
    [ProgId("VBASinc.Connect")]
    public class Connect : IDTExtensibility2
    {
        private AddInHost? _host;

        // Static constructor to catch early type load issues
        static Connect()
        {
            try
            {
                // Force type loading of dependencies
                var _ = typeof(AddInHost);
            }
            catch
            {
                // Silent - if this fails, instance construction will also fail
            }
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                AddInLogger.Log("OnConnection invoked");

                var vbe = application as VBIDE.VBE;
                if (vbe == null)
                {
                    AddInLogger.Log("Direct VBE cast failed, attempting reflection from host application.");
                    try
                    {
                        // Use reflection to get VBE property from hosts like Excel.Application, Word.Application, etc.
                        var vbeProperty = application.GetType().GetProperty("VBE");
                        if (vbeProperty != null)
                        {
                            vbe = vbeProperty.GetValue(application, null) as VBIDE.VBE;
                            AddInLogger.Log("VBE object obtained via reflection.");
                        }
                    }
                    catch (Exception ex)
                    {
                        AddInLogger.Log("Reflection for VBE failed", ex);
                    }
                }

                if (vbe == null)
                {
                    AddInLogger.Log("VBASinc: No se pudo obtener el motor VBE del host.");
                    return;
                }

                _host = new AddInHost(vbe);
                _host.Initialize();

                if (connectMode == ext_ConnectMode.ext_cm_Startup)
                {
                    _host.Start();
                }

                AddInLogger.Log("VBASinc conectado correctamente.");
            }
            catch (Exception ex)
            {
                AddInLogger.Log("Error en OnConnection", ex);
                // DON'T throw - swallow the exception to prevent "can't load" error
                // The add-in will be loaded but non-functional if this happens
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                AddInLogger.Log($"OnDisconnection: {removeMode}");
                _host?.Dispose();
                _host = null;
            }
            catch (Exception ex)
            {
                AddInLogger.Log("Error en OnDisconnection", ex);
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
            try
            {
                AddInLogger.Log("OnStartupComplete");
                _host?.Start();
            }
            catch (Exception ex)
            {
                AddInLogger.Log("Error en OnStartupComplete", ex);
                throw;
            }
        }

        public void OnBeginShutdown(ref Array custom)
        {
            try
            {
                AddInLogger.Log("OnBeginShutdown");
                _host?.Dispose();
                _host = null;
            }
            catch (Exception ex)
            {
                AddInLogger.Log("Error en OnBeginShutdown", ex);
            }
        }
    }
}
