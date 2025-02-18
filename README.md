# 🛠️ VBA Auto Backup for Excel

Este código en **VBA (Visual Basic for Applications)** automatiza la creación de **copias de seguridad** de archivos de Excel cada vez que se abren.

## 🚀 Características
- 📂 Guarda automáticamente **dos copias de seguridad** en ubicaciones predefinidas.
- 🔄 **Refresca todas las conexiones de datos** al abrir el archivo.
- ⚠️ **Manejo de errores** con mensajes en caso de fallo.

## 📌 Instalación
1. **Descarga el archivo `AutoBackup.bas`** de este repositorio.
2. **Importa el módulo en VBA**:
   - Abre Excel y presiona `ALT + F11` para abrir el Editor de VBA.
   - Ve a `Importar archivo` y selecciona `AutoBackup.bas`.
   - Guarda el archivo de Excel como `.xlsm` (Libro habilitado para macros).

## 🏗️ Estructura del Código
- `Workbook_BeforeClose`: Guarda el archivo antes de cerrarlo.
- `Workbook_Open`: Crea dos copias de seguridad con marca de tiempo y refresca los datos.

## 📂 Personalización
Cambia las rutas en las siguientes líneas según tu sistema:

```vba
ruta1 = "C:\Backup_Excel\Seguridad1\" & nombreArchivo
ruta2 = "D:\Copias_Excel\Seguridad2\" & nombreArchivo

## 🔥 ¿Por qué usar esto?
✅ Evitas pérdidas de datos  
✅ Automatizas tareas en Excel  
✅ Mejoras la seguridad de tus archivos  

📩 **¿Tienes mejoras o sugerencias?** ¡Contribuye en el repositorio o abre un `issue`! 🚀
