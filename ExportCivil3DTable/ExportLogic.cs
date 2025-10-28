using System;
using System.IO; 
using System.Text; 

// Referencias de AutoCAD
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices; // Contiene 'Table' de AutoCAD
using Autodesk.AutoCAD.EditorInput; 

// Referencia específica de Civil 3D para Tablas
using Autodesk.Civil.DatabaseServices; // Contiene 'Table' de Civil 3D

namespace ExportCivil3DTable
{
    public class Commands
    {
        [CommandMethod("EXP_TABLA")]
        public static void ExportCivilTableCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n--- Iniciando comando EXP_TABLA ---");

            // 1. Pedir al usuario que seleccione una tabla
            PromptEntityOptions peo = new PromptEntityOptions("\nSeleccione una tabla de Civil 3D:");
            peo.SetRejectMessage("\nEl objeto seleccionado no es una tabla.");

            // 🎯 CORRECCIÓN 1: Ser explícito y usar la Tabla de Civil 3D
            peo.AddAllowedClass(typeof(Autodesk.Civil.DatabaseServices.Table), false); 

            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nComando cancelado.");
                return;
            }

            StringBuilder csvBuilder = new StringBuilder();

            // 2. Procesar la tabla seleccionada
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 🎯 CORRECCIÓN 2: Ser explícito en el tipo de variable y en el 'casting'
                    Autodesk.Civil.DatabaseServices.Table civilTable = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Table;

                    if (civilTable != null)
                    {
                        ed.WriteMessage($"\nLeyendo tabla... {civilTable.Rows.Count} filas encontradas.");

                        // 3. Leer todas las filas y columnas
                        for (int i = 0; i < civilTable.Rows.Count; i++)
                        {
                            string[] rowContents = new string[civilTable.Columns.Count];
                            for (int j = 0; j < civilTable.Columns.Count; j++)
                            {
                                string cellText = civilTable.Cells[i, j].TextString;
                                cellText = cellText.Replace(",", "").Replace("\n", " ");
                                rowContents[j] = cellText;
                            }
                            csvBuilder.AppendLine(string.Join(",", rowContents));
                        }

                        // 4. Guardar el archivo CSV
                        string filePath = @"C:\temp\exportacion.csv";
                        File.WriteAllText(filePath, csvBuilder.ToString());

                        ed.WriteMessage($"\n¡ÉXITO! Tabla exportada a: {filePath}");
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nSe produjo un error durante la exportación: {ex.Message}");
            }
        }
    }
}
