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

            ed.WriteMessage("\n--- Iniciando comando EXP_TABLA (Modo Diagnóstico) ---");

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

                    if (selectedEntity == null)
                    {
                        ed.WriteMessage("\nError: No se pudo leer el objeto seleccionado.");
                        return;
                    }
                    
                    if (selectedEntity is Autodesk.AutoCAD.DatabaseServices.Table)
                    {
                        ed.WriteMessage("\nDetectada tabla estándar de AutoCAD.");
                        tableToRead = selectedEntity as Autodesk.AutoCAD.DatabaseServices.Table;
                    }
                    else if (selectedEntity is Autodesk.Civil.DatabaseServices.Table)
                    {
                        ed.WriteMessage("\nDetectada tabla de Civil 3D. Explotando en memoria...");
                        
                        DBObjectCollection fragments = new DBObjectCollection();
                        selectedEntity.Explode(fragments);

                        ed.WriteMessage($"\n...Análisis de fragmentos: {fragments.Count} objetos encontrados.");
                        
                        // 🎯 CORRECCIÓN: Especificamos el tipo 'DBObject' de AutoCAD
                        foreach (Autodesk.AutoCAD.DatabaseServices.DBObject frag in fragments)
                        {
                            ed.WriteMessage($"\n...Tipo de objeto: {frag.GetType().Name}");
                        }

                        tableToRead = fragments.OfType<Autodesk.AutoCAD.DatabaseServices.Table>().FirstOrDefault();
                    }
                    else
                    {
                        ed.WriteMessage($"\nError: El objeto seleccionado no es una tabla. Tipo de objeto: {selectedEntity.GetType().Name}");
                        tr.Commit();
                        return; 
                    }

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
                        ed.WriteMessage("\nError: El objeto es una tabla de Civil 3D pero no pudo ser convertida a una tabla legible (ningún fragmento era 'Table').");
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
