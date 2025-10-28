using System;
using System.IO; 
using System.Text; 
using System.Linq; 
using System.Collections.Generic; 

// Referencias de AutoCAD
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices; 
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry; 

// Referencia específica de Civil 3D
using Autodesk.Civil.DatabaseServices; 

// Librería de EXCEL
using OfficeOpenXml; 

namespace ExportCivil3DTable
{
    public class Commands
    {
        [CommandMethod("EXP_TABLA")]
        public static void ExportCivilTableCommand()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; 

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n--- Iniciando comando EXP_TABLA (Exportador Excel v3 - Fila Corregida) ---");

            PromptEntityOptions peo = new PromptEntityOptions("\nSeleccione el objeto de tabla a exportar:");
            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK) { /* ... */ }

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.Entity selectedEntity = 
                        tr.GetObject(per.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;

                    if (selectedEntity == null) { /* ... */ }
                    
                    DBObjectCollection fragmentsToRead = new DBObjectCollection();
                    bool isStdTable = false; 

                    if (selectedEntity is Autodesk.AutoCAD.DatabaseServices.Table)
                    {
                        isStdTable = true;
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
                            blockRef.Explode(fragmentsToRead); 
                        }
                    }
                    else
                    {
                        ed.WriteMessage($"\nError: El objeto seleccionado no es una tabla. Tipo: {selectedEntity.GetType().Name}");
                        tr.Commit();
                        return; 
                    }

                    // --- INICIO DE LÓGICA DE EXCEL ---

                    string filePath = @"C:\temp\exportacion.xlsx";
                    FileInfo excelFile = new FileInfo(filePath);
                    if (excelFile.Exists)
                    {
                        excelFile.Delete(); // Borrar el archivo anterior si existe
                    }

                    using (ExcelPackage excelPackage = new ExcelPackage(excelFile))
                    {
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("DatosExportados");
                        ed.WriteMessage("\nCreando archivo Excel...");

                        if (isStdTable)
                        {
                            // ... (Lógica para tablas estándar, sin cambios) ...
                        }
                        else
                        {
                            // --- 🎯 INICIO DE LA LÓGICA CORREGIDA PARA MTEXT 🎯 ---
                            var textObjects = fragmentsToRead.OfType<Autodesk.AutoCAD.DatabaseServices.MText>().ToList();
                            ed.WriteMessage($"\n...Encontrados {textObjects.Count()} objetos MText. Ordenando...");

                            var sortedTexts = textObjects.OrderByDescending(t => t.Location.Y) // De arriba a abajo
                                                         .ThenBy(t => t.Location.X) // De izquierda a derecha
                                                         .ToList();

                            int excelRow = 1;
                            int excelCol = 1;
                            double rowTolerance = 1.0; // Tolerancia para la misma fila (en unidades de dibujo)
                            double currentY = double.MinValue; // Iniciar con un valor que no sea el primero

                            if (sortedTexts.Any())
                            {
                                // Establecer la 'Y' del primer texto como la 'Y' de la primera fila
                                currentY = sortedTexts.First().Location.Y;
                            }

                            foreach (var mtext in sortedTexts)
                            {
                                // Comprobar si este texto está en una nueva fila
                                // (si su 'Y' es significativamente diferente de la 'Y' de la fila actual)
                                if (Math.Abs(mtext.Location.Y - currentY) > rowTolerance)
                                {
                                    // Es una nueva fila
                                    excelRow++;
                                    excelCol = 1; // Reiniciar columna
                                    currentY = mtext.Location.Y; // Actualizar la 'Y' de la fila actual
                                }

                                // Escribir en la celda (excelRow, excelCol)
                                worksheet.Cells[excelRow, excelCol].Value = mtext.Text;
                                
                                // Incrementar la columna para el siguiente texto
                                excelCol++;
                            }
                            // --- 🎯 FIN DE LA LÓGICA CORREGIDA 🎯 ---
                        }

                        excelPackage.Save();
                    }
                    
                    ed.WriteMessage($"\n¡ÉXITO! Datos exportados a: {filePath}");
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nSe produjo un error durante la exportación: {ex.Message}");
                if (ex.InnerException != null)
                {
                    ed.WriteMessage($"\nError interno: {ex.InnerException.Message}");
                }
            }
        }
    }
}
