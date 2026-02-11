<p align="center">
  <img src="https://img.shields.io/badge/version-2.0.2-blue.svg" alt="Version">
  <img src="https://img.shields.io/badge/.NET_Framework-4.7.2+-purple.svg" alt=".NET Framework">
  <img src="https://img.shields.io/badge/Office-2016%2B-green.svg" alt="Office">
  <img src="https://img.shields.io/badge/license-MIT-orange.svg" alt="License">
</p>

<p align="center">
  <img src="logo.png" alt="VBASinc Logo" width="120" />
</p>

<h1 align="center">VBASinc</h1>
<h3 align="center">Motor de Sincronización VBA Bidireccional</h3>

<p align="center">
  Edita código VBA en tu editor favorito (VS Code, Sublime, etc.) y sincroniza automáticamente con Excel, Word, Access y otras aplicaciones Office.
</p>

---

## ✨ Características

| Característica | Descripción |
|----------------|-------------|
| 🔄 **Sincronización bidireccional** | VBA ↔ Archivos externos en tiempo real |
| 👁️ **Detección automática** | FileSystemWatcher + Polling inteligente |
| 📁 **Estructura organizada** | Exporta por carpetas: `Modules/`, `Classes/`, `Forms/` |
| 📊 **Multi-proyecto** | Sincroniza múltiples libros Excel simultáneamente |
| 💾 **Backups automáticos** | Guarda versiones antes de sobrescribir |
| 🎨 **UI moderna** | Panel de control visual con indicadores de estado |

---

## 🚀 Instalación Rápida

### Requisitos
- Windows 10/11
- .NET Framework 4.7.2+
- Microsoft Office 2016+ (32-bit o 64-bit)

### Pasos

1. **Clonar el repositorio**
   ```bash
   git clone https://github.com/tu-usuario/VBASinc.git
   cd VBASinc
   ```

2. **Compilar** (Visual Studio 2019+)
   ```
   Abrir VBASinc.sln → Compilar en Release
   ```

3. **Registrar** (como Administrador)
   ```cmd
   RegistrarComplemento.bat
   ```

4. **Reiniciar Office** y abrir el Editor VBA (`Alt+F11`)

---

## 📖 Uso

### Desde el Editor VBA

1. Abrir VBA con `Alt+F11`
2. Click en **"VBASinc"** en la barra de menú
3. Seleccionar carpeta de exportación
4. Activar **AUTO-SYNC**

### Desde VBA (Programático)

```vba
Sub IniciarSync()
    CreateObject("VBASinc.SyncController").ShowUI ThisWorkbook.VBProject
End Sub

' Con ruta personalizada:
Sub IniciarSyncConRuta()
    CreateObject("VBASinc.SyncController").ShowUI ThisWorkbook.VBProject, "C:\MiProyecto\VBA"
End Sub
```

---

## 📁 Estructura del Proyecto

```
VBASinc/
├── 📄 Connect.cs              # Punto de entrada COM
├── 📄 VBASincSystem.cs        # Interfaz pública VBA
├── 🔧 RegistrarComplemento.bat
│
├── 📂 Host/
│   └── AddInHost.cs           # Controlador principal
│
├── 📂 Sync/
│   ├── SyncEngineV2.cs        # Motor de sincronización
│   ├── ProjectSyncContext.cs  # Contexto multi-proyecto
│   └── ...
│
├── 📂 UI/
│   └── SyncControlForm.cs     # Panel de control
│
└── 📂 docs/
    └── README.md              # Documentación detallada
```

---

## 📂 Archivos Soportados

| Extensión | Tipo | Carpeta |
|-----------|------|---------|
| `.bas` | Módulo Estándar | `Modules/` |
| `.cls` | Clase | `Classes/` |
| `.frm` | Formulario | `Forms/` |

---

## ⚙️ Configuración

Archivo: `%APPDATA%\VBASinc\VBASincSettings.json`

```json
{
  "RootFolderPath": "C:\\src_vba",
  "SyncEnabled": true,
  "PollingIntervalSeconds": 14400,
  "AutoResolveConflicts": false
}
```

---

## 🐛 Solución de Problemas

<details>
<summary><b>El complemento no aparece</b></summary>

1. Ejecutar `RegistrarComplemento.bat` como **Administrador**
2. Verificar claves del registro:
   ```
   HKCU\Software\Microsoft\VBA\VBE\7.1\Addins64\VBASinc.Connect
   ```
3. Reiniciar Office completamente
</details>

<details>
<summary><b>Error "VBProject inválido"</b></summary>

Asegúrate de pasar `ThisWorkbook.VBProject`, no solo `ThisWorkbook`:
```vba
CreateObject("VBASinc.SyncController").ShowUI ThisWorkbook.VBProject
```
</details>

---

## 📜 Changelog

Ver [CHANGELOG.md](CHANGELOG.md) para historial de versiones.

---

## 🤝 Contribuir

1. Fork del repositorio
2. Crear rama: `git checkout -b feature/nueva-funcionalidad`
3. Commit: `git commit -am 'Agregar nueva funcionalidad'`
4. Push: `git push origin feature/nueva-funcionalidad`
5. Crear Pull Request

---

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Ver [LICENSE](LICENSE) para más detalles.

---

<p align="center">
  <b>Desarrollado con ❤️ para la comunidad VBA</b>
</p>
