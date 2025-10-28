using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace ExportCivil3DTable
{
    public class Commands
    {
        // Atributo que registra el comando "EXP_TABLA"
        [CommandMethod("EXP_TABLA")]
        public static void ExportCivilTableCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n--- Comando EXP_TABLA cargado y ejecutado exitosamente ---");
        }
    }
}
