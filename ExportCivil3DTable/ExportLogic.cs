using System;
using System.IO; 
using System.Text; 

// Referencias de AutoCAD
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices; // 👈 Esta es la que usaremos
using Autodesk.AutoCAD.EditorInput; 

// Referencia específica de Civil 3D (la mantenemos por si acaso)
using Autodesk.Civil.DatabaseServices; 

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
            PromptEntityOptions peo = new PromptEntityOptions("\nSeleccione una tabla de AutoCAD/Civil 3D:");
            peo.SetRejectMessage("\nEl objeto seleccionado no es una tabla.");

            // 🎯 CORRECCIÓN LÓGICA 1:
            // Buscamos la tabla de AutoCAD (AcDbTable), no la de Civil (AeccDbTable)
            // Esta es la que tiene las propiedades .Rows, .Columns y .Cells
            peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Table), false); 

            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nComando cancelado.");
                return;
            }

            StringBuilder csvBuilder = new StringBuilder();

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 🎯 CORRECCIÓN LÓGICA 2:
                    // Convertimos el objeto al tipo de tabla de AutoCAD
                    Autodesk.AutoCAD.DatabaseServices.Table acadTable = 
                        tr.GetObject(per.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Table;

                    if (acadTable != null)
                    {
                        // 🎯 CORRECCIÓN LÓGICA 3:
                        // Ahora .Rows, .Columns y .Cells SÍ existen en 'acadTable'
                        ed.WriteMessage($"\nLeyendo tabla... {acadTable.Rows.Count} filas encontradas.");
                        
                        for (int i = 0; i < acadTable.Rows.Count; i++)
                        {
                            string[] rowContents = new string[acadTable.Columns.Count];
                            for (int j = 0; j < acadTable.Columns.Count; j++)
                            {
                                // Usamos la función GetTextString para obtener el valor
                                string cellText = acadTable.Cells[i, j].GetTextString(FormatOption.ForEditing);
                                
                                cellText = cellText.Replace(",", "").Replace("\n", " ");
                                rowContents[j] = cellText;
                            }
                            csvBuilder.AppendLine(string.Join(",", rowContents));
                        }

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
