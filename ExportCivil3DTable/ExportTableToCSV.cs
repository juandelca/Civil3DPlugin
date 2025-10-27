
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
                PromptEntityOptions peo = new PromptEntityOptions("
Selecciona una tabla de Civil 3D:");
                peo.SetRejectMessage("
Debe ser una tabla.");
                peo.AddAllowedClass(typeof(Table), true);
                PromptEntityResult per = ed.GetEntity(peo);

                if (per.Status != PromptStatus.OK) return;

                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    Table table = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Table;

                    if (table == null)
                    {
                        ed.WriteMessage("
No se seleccionó una tabla válida.");
                        return;
                    }

                    string csvPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "TablaCivil3D.csv");
                    using (StreamWriter sw = new StreamWriter(csvPath))
                    {
                        for (int row = 0; row < table.Rows.Count; row++)
                        {
                            string[] cells = new string[table.Columns.Count];
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                cells[col] = table.Cells[row, col].TextString;
                            }
                            sw.WriteLine(string.Join(",", cells));
                        }
                    }

                    ed.WriteMessage($"
Tabla exportada a: {csvPath}");
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"
Error: {ex.Message}");
            }
        }
    }
}
