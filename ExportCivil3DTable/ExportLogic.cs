using System;
using System.IO; 
using System.Text; 
using System.Linq; 

// Referencias de AutoCAD
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices; 
using Autodesk.AutoCAD.EditorInput; 

// Referencia específica de Civil 3D
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

            ed.WriteMessage("\n--- Iniciando comando EXP_TABLA (Doble Explosión) ---");

            PromptEntityOptions peo = new PromptEntityOptions("\nSeleccione el objeto de tabla a exportar:");
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
                    Autodesk.AutoCAD.DatabaseServices.Entity selectedEntity = 
                        tr.GetObject(per.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;

                    if (selectedEntity == null) { /* ... manejo de error ... */ }
                    
                    if (selectedEntity is Autodesk.AutoCAD.DatabaseServices.Table)
                    {
                        ed.WriteMessage("\nDetectada tabla estándar de AutoCAD.");
                        tableToRead = selectedEntity as Autodesk.AutoCAD.DatabaseServices.Table;
                    }
                    else if (selectedEntity is Autodesk.Civil.DatabaseServices.Table)
                    {
                        ed.WriteMessage("\nDetectada tabla de Civil 3D. Explotando (Nivel 1)...");
                        
                        DBObjectCollection fragments_level1 = new DBObjectCollection();
                        selectedEntity.Explode(fragments_level1);
                        ed.WriteMessage($"\n...Nivel 1: {fragments_level1.Count} fragmentos encontrados.");

                        // 🎯 INICIO DE LA NUEVA LÓGICA (DOBLE EXPLOSIÓN) 🎯
                        
                        // Buscar la Referencia a Bloque que encontramos en el paso anterior
                        var blockRef = fragments_level1.OfType<Autodesk.AutoCAD.DatabaseServices.BlockReference>().FirstOrDefault();

                        if (blockRef != null)
                        {
                            ed.WriteMessage("\n...Fragmento es BlockReference. Explotando (Nivel 2)...");
                            
                            DBObjectCollection fragments_level2 = new DBObjectCollection();
                            blockRef.Explode(fragments_level2); // Explotamos el bloque
                            
                            ed.WriteMessage($"\n...Nivel 2: {fragments_level2.Count} fragmentos encontrados.");

                            // Imprimir lo que encontramos en el Nivel 2 (para diagnóstico)
                            foreach (Autodesk.AutoCAD.DatabaseServices.DBObject frag2 in fragments_level2)
                            {
                                ed.WriteMessage($"\n...Tipo de objeto N2: {frag2.GetType().Name}");
                            }

                            // AHORA SÍ, buscamos la tabla en la segunda explosión
                            tableToRead = fragments_level2.OfType<Autodesk.AutoCAD.DatabaseServices.Table>().FirstOrDefault();
                        }
                        else
                        {
                            ed.WriteMessage("\n...Error: La explosión de Nivel 1 no contenía un BlockReference.");
                        }
                        // 🎯 FIN DE LA NUEVA LÓGICA 🎯
                    }
                    else
                    {
                        ed.WriteMessage($"\nError: El objeto seleccionado no es una tabla. Tipo: {selectedEntity.GetType().Name}");
                        tr.Commit();
                        return; 
                    }

                    // --- INICIO DE LECTURA DE TABLA (Sin cambios) ---
                    if (tableToRead != null)
                    {
                        ed.WriteMessage($"\n¡Tabla encontrada! Leyendo... {tableToRead.Rows.Count} filas.");
                        
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
                        ed.WriteMessage("\nError: La explosión final no contenía un objeto 'Table' legible.");
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
