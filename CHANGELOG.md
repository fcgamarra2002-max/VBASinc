# Changelog

Todos los cambios notables en este proyecto serán documentados aquí.

El formato está basado en [Keep a Changelog](https://keepachangelog.com/es/1.0.0/),
y este proyecto adhiere a [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.0.2] - 2026-02-08

### ✨ Añadido
- **Soporte multi-proyecto**: Sincroniza múltiples libros Excel simultáneamente
- **Indicador de estado SVG**: Círculos rojo/verde para mostrar estado de sincronización
- **Selector de proyectos**: ComboBox para cambiar entre proyectos abiertos
- **Botón refrescar proyectos**: Detecta nuevos libros abiertos dinámicamente
- **Estructura de carpetas VBA**: Exporta a `Modules/`, `Classes/` y `Forms/`
- **Iconos SVG programáticos**: Flechas de exportar/importar y refrescar

### 🔧 Cambiado
- Título del panel: "MOTOR DE SINCRONIZACIÓN VBA" con versión centrada
- Formulario con tamaño fijo (no redimensionable)
- Botón AUTO-SYNC ahora trabaja por proyecto individual

### 🐛 Corregido
- Compatibilidad con C# 8.0 (removido patrones `or`)
- Posición del botón Cerrar dentro del área visible

---

## [2.0.1] - 2026-02-05

### ✨ Añadido
- TreeView jerárquico para selección de módulos
- Historial de cambios visual con íconos y colores
- Contadores de módulos internos/externos

### 🔧 Cambiado
- Mejoras en la interfaz de usuario
- Reorganización de controles

---

## [2.0.0] - 2026-02-01

### ✨ Añadido
- Motor de sincronización V2 completo
- Panel de control moderno
- Sincronización bidireccional automática
- Detección de conflictos
- Sistema de backups

### 🔧 Cambiado
- Arquitectura completamente rediseñada
- Nueva UI con diseño moderno

---

## [1.0.0] - 2026-01-15

### ✨ Añadido
- Versión inicial
- Exportación/Importación básica
- Registro COM para VBE 6.0 y 7.1
