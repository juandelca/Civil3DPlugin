using System;
using System.IO; // Necesario para escribir archivos
using System.Text; // Necesario para construir el texto CSV

// Referencias de AutoCAD
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput; // Necesario para que el usuario seleccione cosas

// Referencia específica de Civil 3D para Tablas
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
            PromptEntityOptions peo = new PromptEntityOptions("\nSeleccione una tabla de Civil 3D:");
            peo.SetRejectMessage("\nEl objeto seleccionado no es una tabla.");
            peo.AddAllowedClass(typeof(Autodesk.Civil.DatabaseServices.Table), false); // Solo permite seleccionar Tablas de Civil 3D

            PromptEntityResult per = ed.GetEntity(peo);

            // Si el usuario no selecciona nada o cancela
            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nComando cancelado.");
                return;
            }

            // Variable para guardar el texto del CSV
            StringBuilder csvBuilder = new StringBuilder();

            // 2. Procesar la tabla seleccionada
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Obtener la tabla que el usuario seleccionó
                    Table civilTable = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Table;

                    if (civilTable != null)
                    {
                        ed.WriteMessage($"\nLeyendo tabla... {civilTable.Rows.Count} filas encontradas.");

                        // 3. Leer todas las filas y columnas
                        for (int i = 0; i < civilTable.Rows.Count; i++)
                        {
                            string[] rowContents = new string[civilTable.Columns.Count];
                            for (int j = 0; j < civilTable.Columns.Count; j++)
                            {
                                // Obtener el texto de la celda
                                string cellText = civilTable.Cells[i, j].TextString;
                                
                                // Limpiar el texto (remover comas y saltos de línea para que no rompan el CSV)
                                cellText = cellText.Replace(",", "").Replace("\n", " ");
                                
                                rowContents[j] = cellText;
                            }
                            // Unir todas las celdas de esta fila con una coma
                            csvBuilder.AppendLine(string.Join(",", rowContents));
                        }

                        // 4. Guardar el archivo CSV
                        // 
                        //  ¡¡¡ IMPORTANTE !!!
                        //  Crea una carpeta llamada "temp" en tu disco C:
                        //  O cambia esta ruta a una carpeta que exista.
                        //
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
