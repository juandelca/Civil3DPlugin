using System;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;

[assembly: CommandClass(typeof(ExportCivil3DTable.ExportTableToCSV))]

namespace ExportCivil3DTable
{
    public class ExportTableToCSV
    {
        [CommandMethod("EXPORTTABLECSV")]
        public void ExportTableToCSVCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // CORRECCIÓN LÍNEA 24: Se utiliza "\n" dentro de la cadena (o se utiliza una cadena verbatim @)
                // Usaremos "\n" que es el estándar para salto de línea.
                PromptEntityOptions peo = new PromptEntityOptions("\nSelecciona una tabla de Civil 3D:");
                
                // CORRECCIÓN LÍNEA 25: Se utiliza "\n"
                peo.SetRejectMessage("\nDebe ser una tabla.");
                peo.AddAllowedClass(typeof(Table), true);
                PromptEntityResult per = ed.GetEntity(peo);

                if (per.Status != PromptStatus.OK) return;

                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    Table table = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Table;

                    if (table == null)
                    {
                        // CORRECCIÓN LÍNEA 33: Se utiliza "\n"
                        ed.WriteMessage("\nNo se seleccionó una tabla válida.");
                        return;
                    }

                    // Se recomienda usar el prefijo @ para evitar problemas con las barras invertidas en las rutas
                    string csvPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "TablaCivil3D.csv");
                    
                    using (StreamWriter sw = new StreamWriter(csvPath))
                    {
                        for (int row = 0; row < table.Rows.Count; row++)
                        {
                            string[] cells = new string[table.Columns.Count];
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                // Se añade un reemplazo de coma, si existe en el texto, para no romper el CSV
                                string cellText = table.Cells[row, col].TextString;
                                // Envuelve el texto en comillas si contiene comas
                                if (cellText.Contains(","))
                                {
                                    cells[col] = $"\"{cellText.Replace("\"", "\"\"")}\"";
                                }
                                else
                                {
                                    cells[col] = cellText;
                                }
                            }
                            sw.WriteLine(string.Join(",", cells));
                        }
                    }

                    // CORRECCIÓN LÍNEA 49: Se utiliza "\n"
                    ed.WriteMessage($"\nTabla exportada a: {csvPath}");
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                // CORRECCIÓN LÍNEA 53: Se utiliza "\n"
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }
    }
}
