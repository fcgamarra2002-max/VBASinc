# 🚀 VBASinc v1.0.4 - Professional VBA Sync Engine

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Version](https://img.shields.io/badge/Version-1.0.4-blue.svg)]()
[![Platform: Office](https://img.shields.io/badge/Platform-Office%20%2F%20VBA-orange.svg)]()

**VBASinc** is a high-performance, event-driven synchronization engine designed to bridge the gap between Microsoft Office VBA (Excel, Word, PowerPoint) and modern version control systems (Git, SVN, etc.). It allows developers to export/import VBA modules automatically and in real-time, providing a seamless DevOps experience for legacy Office environments.

---

## 📥 [Download Latest Installer (VBASinc.exe)](https://github.com/fcgamarra2002-max/VBASinc/raw/main/VBASinc.exe)

> **Note**: This is the all-in-one professional installer v1.0.4.1. Run as Administrator to automatically configure Office security and register the engine.

---

## 🌟 Key Features

*   **Real-Time Reactive Sync**: Powered by `FileSystemWatcher` and VBA IDE events. No polling, no lag.
*   **Zero-Configuration Installer**: A self-contained `.exe` that handles COM registration and Office Security automatically.
*   **AppData Centric**: The core engine lives in `%AppData%`, keeping your root drive clean and professional.
*   **Security Automation**: Automatically configures "Trusted Locations" and "Trust Access to VBA Object Model" for a hassle-free experience.
*   **Multi-App Support**: Compatible with Excel, Word, and PowerPoint (Office 2016 to Office 365).
*   **Bilingual Logs**: Detailed installation and synchronization logs in both Spanish and English.

---

## 🛠️ Installation & Usage

1.  **Download**: Get the latest `VBASinc.exe` from this repository.
2.  **Run**: Execute `VBASinc.exe` as Administrator.
    *   *What happens?* It extracts the engine to AppData, registers the COM component, and clears Office security blockages.
3.  **Launch**: Open the VBA Editor (Alt+F11) in Excel, Word, or PowerPoint.
4.  **Sync**: Find the **VBASinc** menu and start managing your modules like a pro.

---

## 📦 Project Structure

*   `VBASinc.dll`: The core COM-AddIn engine (Sync Logic).
*   `SetupInstaller.cs`: The professional "All-in-One" installer code.
*   `UI/`: Modern Windows Forms interface for configuration and monitoring.
*   `Host/`: Interop layer to communicate with the VBA IDE.

---

## 👨‍💻 Developer / Autor

**fcgamarra2002-max**

Developed with a focus on efficiency and reliability for modern VBA development workflows.

---

# 🇪🇸 Versión en Español

**VBASinc** es un motor de sincronización de alto rendimiento diseñado para conectar VBA (Excel, Word, PowerPoint) con sistemas de control de versiones.

### Características Principales:
- **Sincronización Reactiva**: Sin esperas, detecta cambios al instante.
- **Instalador Todo-en-Uno**: Configura la seguridad de Office y registra la DLL automáticamente.
- **Limpio y Portable**: Se instala en `AppData` para no generar archivos basura en la raíz del disco.

## Licencia / License
Distributed under the **MIT License**. See `LICENSE` for more information.

---
*Created with ❤️ for the VBA Community.*
