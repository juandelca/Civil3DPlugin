using System;
using System.IO; 
using System.Text; 
using System.Linq; // Necesario para buscar en la colección de fragmentos

// Referencias de AutoCAD
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices; // La tabla estándar
using Autodesk.AutoCAD.EditorInput; 

// Referencia específica de Civil 3D
using Autodesk.Civil.DatabaseServices; // La tabla de Civil 3D

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

            // 1. Pedir al usuario que seleccione CUALQUIER tipo de tabla
            PromptEntityOptions peo = new PromptEntityOptions("\nSeleccione una tabla de AutoCAD o Civil 3D:");
            peo.SetRejectMessage("\nEl objeto seleccionado no es una tabla.");
            
            // 🎯 CORRECCIÓN: Aceptamos ambos tipos de tabla
            peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Table), true);
            peo.AddAllowedClass(typeof(Autodesk.Civil.DatabaseServices.Table), true);

            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nComando cancelado.");
                return;
            }

            // Variable para la tabla estándar de AutoCAD, que es la que leeremos
            Autodesk.AutoCAD.DatabaseServices.Table tableToRead = null;
            StringBuilder csvBuilder = new StringBuilder();

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Abrir el objeto seleccionado
                    Entity selectedEntity = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;

                    // 2. Comprobar qué tipo de tabla es
                    
                    // Opción A: Es una tabla estándar de AutoCAD
                    if (selectedEntity is Autodesk.AutoCAD.DatabaseServices.Table)
                    {
                        ed.WriteMessage("\nDetectada tabla estándar de AutoCAD.");
                        tableToRead = selectedEntity as Autodesk.AutoCAD.DatabaseServices.Table;
                    }
                    // Opción B: Es una tabla de Civil 3D (como tu AECC_LEGEND_TABLE)
                    else if (selectedEntity is Autodesk.Civil.DatabaseServices.Table)
                    {
                        ed.WriteMessage("\nDetectada tabla de Civil 3D. Explotando en memoria...");
                        
                        // Explotarla en memoria para convertirla a una tabla de AutoCAD
                        DBObjectCollection fragments = new DBObjectCollection();
                        selectedEntity.Explode(fragments);

                        // Buscar la tabla estándar de AutoCAD entre los fragmentos
                        tableToRead = fragments.OfType<Autodesk.AutoCAD.DatabaseServices.Table>().FirstOrDefault();
                    }

                    // 3. Procesar la tabla (si la encontramos)
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
                        ed.WriteMessage("\nError: No se pudo convertir la tabla de Civil 3D a una tabla legible.");
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
