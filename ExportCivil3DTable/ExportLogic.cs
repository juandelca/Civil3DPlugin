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

            ed.WriteMessage("\n--- Iniciando comando EXP_TABLA (Lector de MText) ---");

            PromptEntityOptions peo = new PromptEntityOptions("\nSeleccione el objeto de tabla a exportar:");
            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nComando cancelado.");
                return;
            }

            // Ya no buscamos un 'Table', así que 'tableToRead' se elimina.
            StringBuilder csvBuilder = new StringBuilder();

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.Entity selectedEntity = 
                        tr.GetObject(per.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;

                    if (selectedEntity == null) { /* ... */ }
                    
                    // Colección final de fragmentos a leer
                    DBObjectCollection fragmentsToRead = new DBObjectCollection();

                    if (selectedEntity is Autodesk.AutoCAD.DatabaseServices.Table)
                    {
                        ed.WriteMessage("\nDetectada tabla estándar de AutoCAD. Leyendo celdas...");
                        
                        // Si es una tabla estándar, la leemos como antes
                        Autodesk.AutoCAD.DatabaseServices.Table stdTable = selectedEntity as Autodesk.AutoCAD.DatabaseServices.Table;
                        for (int i = 0; i < stdTable.Rows.Count; i++)
                        {
                            string[] rowContents = new string[stdTable.Columns.Count];
                            for (int j = 0; j < stdTable.Columns.Count; j++)
                            {
                                rowContents[j] = stdTable.Cells[i, j].GetTextString(FormatOption.ForEditing).Replace(",", "").Replace("\n", " ");
                            }
                            csvBuilder.AppendLine(string.Join(",", rowContents));
                        }
                    }
                    else if (selectedEntity is Autodesk.Civil.DatabaseServices.Table)
                    {
                        ed.WriteMessage("\nDetectada tabla de Civil 3D. Explotando (Nivel 1)...");
                        
                        DBObjectCollection fragments_level1 = new DBObjectCollection();
                        selectedEntity.Explode(fragments_level1);

                        var blockRef = fragments_level1.OfType<Autodesk.AutoCAD.DatabaseServices.BlockReference>().FirstOrDefault();

                        if (blockRef != null)
                        {
                            ed.WriteMessage("\n...Nivel 1 es BlockReference. Explotando (Nivel 2)...");
                            blockRef.Explode(fragmentsToRead); // Explotamos el bloque y guardamos los fragmentos
                        }
                        else
                        {
                             ed.WriteMessage("\n...Advertencia: La explosión de Nivel 1 no contenía un BlockReference. Usando fragmentos de Nivel 1.");
                             fragmentsToRead = fragments_level1;
                        }
                    }
                    else
                    {
                        ed.WriteMessage($"\nError: El objeto seleccionado no es una tabla. Tipo: {selectedEntity.GetType().Name}");
                        tr.Commit();
                        return; 
                    }


                    // 🎯 INICIO DE LA NUEVA LÓGICA: LEER MTEXT 🎯

                    // Si NO es una tabla estándar, leemos los fragmentos (MText)
                    if (!(selectedEntity is Autodesk.AutoCAD.DatabaseServices.Table))
                    {
                        // Filtramos solo los objetos MText de los fragmentos
                        var textObjects = fragmentsToRead.OfType<Autodesk.AutoCAD.DatabaseServices.MText>();
                        
                        int textCount = 0;
                        if (textObjects.Any())
                        {
                            ed.WriteMessage($"\n...Encontrados {textObjects.Count()} objetos MText. Leyendo...");
                            foreach (var mtext in textObjects)
                            {
                                string cellText = mtext.TextString; // Leemos el texto del objeto MText
                                cellText = cellText.Replace(",", "").Replace("\n", " ");
                                csvBuilder.AppendLine(cellText); // Escribimos cada texto en una nueva línea del CSV
                                textCount++;
                            }
                        }
                        else
                        {
                            ed.WriteMessage("\nError: La explosión no contenía ningún objeto MText legible.");
                            tr.Commit();
                            return;
                        }
                    }
                    
                    // --- GUARDAR EL ARCHIVO (Sin cambios) ---
                    string filePath = @"C:\temp\exportacion.csv";
                    File.WriteAllText(filePath, csvBuilder.ToString());

                    ed.WriteMessage($"\n¡ÉXITO! Datos exportados a: {filePath}");
                    
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
