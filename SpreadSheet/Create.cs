using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace SpreadSheet
{
    public class Create
    {
        /// <summary>
        /// Creates a spreadsheet document with a specified filepath.
        /// </summary>
        /// <param name="filepath"> Complete path</param>
        public static void CreateSpreadsheetWorkbook(string filepath)
        {
            //if (UtilityFunctions.IsFileOpen(filepath) && )
            //{
            //    MessageBox.Show("File is open. Please close it before creating a new file.");
            //}
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
                sheets.Append(sheet);
            }
        }
        public static void SaveToNewFileIfOpen(string filepath)
        {
            if (UtilityFunctions.IsFileOpen(filepath))
            {
                string newFilePath = Path.Combine(Path.GetDirectoryName(filepath), "Backup_" + Path.GetFileName(filepath));
                Console.WriteLine($"File is open. Saving to: {newFilePath}");
                CreateSpreadsheetWorkbook(newFilePath);
            }
            else
            {
                CreateSpreadsheetWorkbook(filepath);
            }
        }

        /// <summary>
        /// Inserts text into a specified sheet in a spreadsheet document.
        /// If sheet doesnot exists it will create new sheet with the name
        /// If the file does not exists it will create new file
        /// </summary>
        /// <param name="docName">File name along with path</param>
        /// <param name="inputData">List of list were each inner list is a row</param>
        /// <param name="sheetName">Name of sheet</param>
        public static void InsertData(string docName, List<List<string>> inputData, string sheetName)
        {
            if (!File.Exists(docName))
            {
                CreateSpreadsheetWorkbook(docName);
            }
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                WorkbookPart workbookPart = spreadSheet.WorkbookPart ?? spreadSheet.AddWorkbookPart();

                // Get the SharedStringTablePart. If it does not exist, create a new one.
                SharedStringTablePart shareStringPart;
                if (workbookPart.GetPartsOfType<SharedStringTablePart>().Any())
                {
                    shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                }
                // If sheetName is null, use the first sheet or create a new one
                WorksheetPart worksheetPart = string.IsNullOrEmpty(sheetName)
                    ? UtilityFunctions.GetFirstWorksheetPart(workbookPart) // Get the first sheet or create a new one
                    : UtilityFunctions.GetWorksheetPartByName(workbookPart, sheetName); // Get by name if provided

                // If no sheet is found, create a new one with a default name
                if (worksheetPart == null)
                {
                    string newSheetName = string.IsNullOrEmpty(sheetName) ? "Sheet1" : sheetName;
                    worksheetPart = UtilityFunctions.InsertWorksheet(workbookPart, newSheetName); // Assuming InsertWorksheet can handle the sheet name
                }

                //// Find the last used row index to start appending
                //int lastRowIndex = UtilityFunctions.GetLastRowIndex(worksheetPart);
                //int rowIndex = lastRowIndex + 1; // Start writing after the last row

                //// Insert a new worksheet.
                //WorksheetPart worksheetPart = UtilityFunctions.InsertWorksheet(workbookPart);

                for (int i = 0; i < inputData.Count; i++) // Rows
                {
                    for (int j = 0; j < inputData[i].Count; j++) // Columns
                    {
                        // Example: 0 => "A", 1 => "B", etc.
                        string columnName = UtilityFunctions.GetColumnNameFromIndex(j);
                        uint rowIndex = (uint)(i + 1); // Excel rows are 1-based

                        Cell cell = UtilityFunctions.InsertCellInWorksheet(columnName, rowIndex, worksheetPart);

                        string value = inputData[i][j];

                        if (double.TryParse(value, out double numericValue))
                        {
                            // Value is a number
                            cell.CellValue = new CellValue(numericValue.ToString(System.Globalization.CultureInfo.InvariantCulture));
                            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        else
                        {
                            // Value is text
                            int index = UtilityFunctions.InsertSharedStringItem(value, shareStringPart);
                            cell.CellValue = new CellValue(index.ToString());
                            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                        }
                    }
                }
            }
        }
        public static void CalculateSumOfCellRange(string docName, string worksheetName, string firstCellName, string lastCellName, string resultCell)
        {
            // Open the document for editing.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
            {
                IEnumerable<Sheet> sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName) ?? Enumerable.Empty<Sheet>();
                string firstId = sheets.FirstOrDefault()?.Id?.Value ?? string.Empty; // Adjusted to avoid nullable reference types
                if (!sheets.Any() || string.IsNullOrEmpty(firstId))
                {
                    // The specified worksheet does not exist.
                    return;
                }

                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(firstId);
                Worksheet worksheet = worksheetPart.Worksheet;

                // Get the row number and column name for the first and last cells in the range.
                uint firstRowNum = UtilityFunctions.GetRowIndex(firstCellName);
                uint lastRowNum = UtilityFunctions.GetRowIndex(lastCellName);
                string firstColumn = UtilityFunctions.GetColumnNameFromCellReference(firstCellName);
                string lastColumn = UtilityFunctions.GetColumnNameFromCellReference(lastCellName);

                double sum = 0;

                // Iterate through the cells within the range and add their values to the sum.
                foreach (Row row in worksheet.Descendants<Row>().Where(r => r.RowIndex != null && r.RowIndex.Value >= firstRowNum && r.RowIndex.Value <= lastRowNum))
                {
                    foreach (Cell cell in row)
                    {
                        if (cell.CellReference != null && cell.CellReference.Value != null)
                        {
                            string columnName = UtilityFunctions.GetColumnNameFromCellReference(cell.CellReference.Value);
                            if (UtilityFunctions.CompareColumn(columnName, firstColumn) >= 0 && UtilityFunctions.CompareColumn(columnName, lastColumn) <= 0 && double.TryParse(cell.CellValue?.Text, out double num))
                            {
                                sum += num;
                            }
                        }
                    }
                }

                // Get the SharedStringTablePart and add the result to it.
                // If the SharedStringPart does not exist, create a new one.
                SharedStringTablePart shareStringPart;
                if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }

                // Insert the result into the SharedStringTablePart.
                int index = UtilityFunctions.InsertSharedStringItem("Result: " + sum, shareStringPart);

                Cell result = UtilityFunctions.InsertCellInWorksheet(UtilityFunctions.GetColumnNameFromCellReference(resultCell), UtilityFunctions.GetRowIndex(resultCell), worksheetPart);

                // Set the value of the cell.
                result.CellValue = new CellValue(index.ToString());
                result.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            }
        }
    }
}
