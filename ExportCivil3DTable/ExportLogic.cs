using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace ExportCivil3DTable
{
    public class Commands
    {
        // 🚨 Este atributo registra el comando que debes teclear en la línea de comandos de Civil 3D. 🚨
        [CommandMethod("EXP_TABLA")]
        public static void ExportCivilTableCommand()
        {
            // Obtener el documento activo para interactuar con la interfaz
            Document doc = Application.DocumentManager.MdiActiveDocument;
            
            // Muestra un mensaje para confirmar que el comando se ejecutó
            doc.Editor.WriteMessage("\n--- Comando EXP_TABLA cargado y ejecutado exitosamente ---");

            // *** Tu lógica de programación para trabajar con las tablas de Civil 3D va aquí ***

            doc.Editor.WriteMessage("\n¡El plugin ha terminado de ejecutarse!");
        }
    }
}
