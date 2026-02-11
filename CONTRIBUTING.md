# Contribuir a VBASinc

¡Gracias por tu interés en contribuir! 🎉

## 🚀 Cómo Contribuir

### Reportar Bugs

1. Verificar que el bug no haya sido reportado previamente
2. Crear un issue con:
   - Descripción clara del problema
   - Pasos para reproducir
   - Comportamiento esperado vs actual
   - Versión de Office y Windows

### Proponer Mejoras

1. Abrir un issue describiendo la mejora
2. Esperar feedback antes de implementar
3. Seguir el proceso de Pull Request

### Pull Requests

1. Fork del repositorio
2. Crear rama desde `main`:
   ```bash
   git checkout -b feature/mi-mejora
   ```
3. Hacer cambios siguiendo las convenciones
4. Commit con mensaje descriptivo:
   ```bash
   git commit -m "feat: agregar soporte para módulos de documento"
   ```
5. Push y crear Pull Request

## 📝 Convenciones

### Commits
Seguimos [Conventional Commits](https://www.conventionalcommits.org/):
- `feat:` nueva funcionalidad
- `fix:` corrección de bug
- `docs:` documentación
- `refactor:` refactorización
- `test:` tests

### Código
- Usar nomenclatura C# estándar (PascalCase para métodos/clases)
- Comentarios en español
- Documentar métodos públicos con XML docs

## 🛠️ Desarrollo Local

1. Abrir `VBASinc.sln` en Visual Studio
2. Compilar en Release
3. Ejecutar `RegistrarComplemento.bat` como Admin
4. Probar en Excel/Word

## 📄 Licencia

Al contribuir, aceptas que tus contribuciones se licencien bajo MIT.
