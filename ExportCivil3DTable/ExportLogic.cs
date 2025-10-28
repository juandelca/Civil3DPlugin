using System;
using System.IO; 
using System.Text; 
using System.Linq; 
using System.Collections.Generic; // Para listas

// Referencias de AutoCAD
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices; 
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry; // Para Point3d (coordenadas)

// Referencia específica de Civil 3D
using Autodesk.Civil.DatabaseServices; 

// 🎯 NUEVA LIBRERÍA DE EXCEL 🎯
using OfficeOpenXml; 

namespace ExportCivil3DTable
{
    public class Commands
    {
        [CommandMethod("EXP_TABLA")]
        public static void ExportCivilTableCommand()
        {
            // --- Configuración de Licencia de EPPlus (requerido) ---
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // O 'LicenseContext.Commercial' si tienes licencia

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n--- Iniciando comando EXP_TABLA (Exportador Excel v2) ---");

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
                    bool isStdTable = false; // Flag para saber si es tabla estándar

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

                    // --- 🎯 INICIO DE LA NUEVA LÓGICA DE EXCEL 🎯 ---

                    // 1. Definir la ruta del archivo Excel
                    string filePath = @"C:\temp\exportacion.xlsx";
                    FileInfo excelFile = new FileInfo(filePath);

                    // 2. Crear el paquete Excel
                    using (ExcelPackage excelPackage = new ExcelPackage(excelFile))
                    {
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("DatosExportados");
                        ed.WriteMessage("\nCreando archivo Excel...");

                        // --- Opción A: Es una tabla estándar (fácil) ---
                        if (isStdTable)
                        {
                            Autodesk.AutoCAD.DatabaseServices.Table stdTable = selectedEntity as Autodesk.AutoCAD.DatabaseServices.Table;
                            ed.WriteMessage("\nEscribiendo tabla estándar en Excel...");
                            for (int i = 0; i < stdTable.Rows.Count; i++)
                            {
                                for (int j = 0; j < stdTable.Columns.Count; j++)
                                {
                                    string cellText = stdTable.Cells[i, j].GetTextString(FormatOption.ForEditing);
                                    // Escribir en Excel (filas/columnas +1 porque Excel empieza en 1)
                                    worksheet.Cells[i + 1, j + 1].Value = cellText;
                                }
                            }
                        }
                        // --- Opción B: Es un montón de MText (difícil, requiere ordenar) ---
                        else
                        {
                            var textObjects = fragmentsToRead.OfType<Autodesk.AutoCAD.DatabaseServices.MText>().ToList();
                            ed.WriteMessage($"\n...Encontrados {textObjects.Count()} objetos MText. Ordenando por coordenadas...");

                            // 3. Ordenar los MText por posición Y (de arriba a abajo) y luego X (de izquierda a derecha)
                            // En AutoCAD, Y crece hacia arriba, así que ordenamos Y de forma DESCENDENTE.
                            var sortedTexts = textObjects.OrderByDescending(t => t.Position.Y).ThenBy(t => t.Position.X);

                            // 4. Escribir los textos ordenados en el Excel
                            int excelRow = 1;
                            double currentY = double.MaxValue;

                            foreach (var mtext in sortedTexts)
                            {
                                // Si la coordenada Y es muy diferente, asumimos que es una nueva fila
                                // (Ajustar este '1.0' si el espaciado es diferente)
                                if (Math.Abs(mtext.Position.Y - currentY) > 1.0 && excelRow > 1)
                                {
                                    excelRow++;
                                }
                                
                                // Escribimos el texto en la primera columna disponible de esa fila
                                int excelCol = 1;
                                while (worksheet.Cells[excelRow, excelCol].Value != null)
                                {
                                    excelCol++;
                                }
                                worksheet.Cells[excelRow, excelCol].Value = mtext.Text;

                                currentY = mtext.Position.Y; // Guardar la coordenada Y de esta fila
                            }
                        }

                        // 5. Guardar el archivo Excel
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
