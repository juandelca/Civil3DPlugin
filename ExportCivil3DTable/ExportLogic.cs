using System;
using System.IO; 
using System.Text; 
using System.Linq; 

// Referencias de AutoCAD
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices; // Contiene la 'Entity' base y la 'Table' de AutoCAD
using Autodesk.AutoCAD.EditorInput; 

// Referencia específica de Civil 3D
using Autodesk.Civil.DatabaseServices; // Contiene la 'Entity' y 'Table' de Civil 3D

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

            // 1. 🎯 CORRECCIÓN: Quitamos el filtro
            // Dejamos que el usuario seleccione CUALQUIER entidad.
            PromptEntityOptions peo = new PromptEntityOptions("\nSeleccione el objeto de tabla a exportar:");
            // Ya no hay 'AddAllowedClass' ni 'SetRejectMessage'

            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nComando cancelado.");
                return;
            }

            Autodesk.AutoCAD.DatabaseServices.Table tableToRead = null;
            StringBuilder csvBuilder = new StringBuilder();

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Abrimos la entidad seleccionada
                    Autodesk.AutoCAD.DatabaseServices.Entity selectedEntity = 
                        tr.GetObject(per.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;

                    if (selectedEntity == null)
                    {
                        ed.WriteMessage("\nError: No se pudo leer el objeto seleccionado.");
                        return;
                    }

                    // 2. 🎯 CORRECCIÓN: Analizamos el objeto DESPUÉS de seleccionarlo
                    
                    // Opción A: Es una tabla estándar de AutoCAD
                    if (selectedEntity is Autodesk.AutoCAD.DatabaseServices.Table)
                    {
                        ed.WriteMessage("\nDetectada tabla estándar de AutoCAD.");
                        tableToRead = selectedEntity as Autodesk.AutoCAD.DatabaseServices.Table;
                    }
                    // Opción B: Es una tabla de Civil 3D (como AECC_LEGEND_TABLE)
                    else if (selectedEntity is Autodesk.Civil.DatabaseServices.Table)
                    {
                        ed.WriteMessage("\nDetectada tabla de Civil 3D. Explotando en memoria...");
                        
                        DBObjectCollection fragments = new DBObjectCollection();
                        selectedEntity.Explode(fragments);

                        tableToRead = fragments.OfType<Autodesk.AutoCAD.DatabaseServices.Table>().FirstOrDefault();
                    }
                    // Opción C: Es cualquier otra cosa
                    else
                    {
                        ed.WriteMessage($"\nError: El objeto seleccionado no es una tabla. Tipo de objeto: {selectedEntity.GetType().Name}");
                        tr.Commit();
                        return; // Salir del comando
                    }

                    // 3. Procesar la tabla (si se encontró una)
                    if (tableToRead != null)
                    {
                        ed.WriteMessage($"\nLeyendo tabla... {tableToRead.Rows.Count} filas encontradas.");
                        
                        for (int i = 0; i < tableToRead.Rows.Count; i++)
                        {
                            string[] rowContents = new string[tableToRead.Columns.Count];
                            for (int j = 0; j < tableToRead.Columns.Count; j++)
                            {
                                string cellText = tableToRead.Cells[i, j].GetTextString(FormatOption.ForEditing);
                                cellText = cellText.Replace(",", "").Replace("\n", " ");
                                rowContents[j] = cellText;
                            }
                            csvBuilder.AppendLine(string.Join(",", rowContents));
                        }

                        string filePath = @"C:\temp\exportacion.csv";
                        File.WriteAllText(filePath, csvBuilder.ToString());

                        ed.WriteMessage($"\n¡ÉXITO! Tabla exportada a: {filePath}");
                    }
                    else
                    {
                        ed.WriteMessage("\nError: El objeto es una tabla de Civil 3D pero no pudo ser convertida a una tabla legible.");
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
