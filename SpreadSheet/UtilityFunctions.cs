using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SpreadSheet
{
    public class UtilityFunctions
    {
        //Extracts the numeric row index from a cell reference (e.g., "B3" → 3).
        public static uint GetRowIndex(string cellReference)
        {
            string rowPart = new string(cellReference.Where(c => char.IsDigit(c)).ToArray());
            return uint.TryParse(rowPart, out uint rowIndex) ? rowIndex : 0;
        }

        //Extracts the column name from a cell reference(e.g., "B3" → "B").
        public static string GetColumnNameFromCellReference(string cellReference)
        {
            return new string(cellReference.Where(c => char.IsLetter(c)).ToArray());
        }

        //Converts a column index to its corresponding Excel column name (e.g., 0 → "A", 1 → "B", ..., 25 → "Z", 26 → "AA").
        public static string GetColumnNameFromIndex(int columnIndex)
        {
            int dividend = columnIndex + 1;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }


        //Compares column letters(e.g., "A" < "B", "Z" < "AA").
        public static int CompareColumn(string columnA, string columnB)
        {
            int GetColumnIndex(string col)
            {
                int index = 0;
                foreach (char c in col.ToUpper())
                {
                    index *= 26;
                    index += (c - 'A' + 1);
                }
                return index;
            }

            return GetColumnIndex(columnA).CompareTo(GetColumnIndex(columnB));
        }

        ////Adds a string to the shared string table(or returns its existing index).
        //// Helper function to insert text into the SharedStringTable.
        //public static int InsertSharedStringItem<T>(T text, SharedStringTablePart shareStringPart)
        //{
        //    if (shareStringPart.SharedStringTable == null)
        //    {
        //        shareStringPart.SharedStringTable = new SharedStringTable();
        //    }

        //    int i = 0;
        //    foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        //    {
        //        if (item.InnerText == text.ToString())
        //        {
        //            return i;
        //        }
        //        i++;
        //    }

        //    shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text.ToString())));
        //    shareStringPart.SharedStringTable.Save();

        //    return i;
        //}

        private static Dictionary<string, int> sharedStringLookup = new Dictionary<string, int>();
        public static int InsertSharedStringItem<T>(T text, SharedStringTablePart shareStringPart)
        {
            string value = text.ToString();

            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            // Initialize lookup dictionary if it's empty
            if (sharedStringLookup.Count == 0)
            {
                int index = 0;
                foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
                {
                    sharedStringLookup[item.InnerText] = index;
                    index++;
                }
            }

            // Check if the value already exists
            if (sharedStringLookup.TryGetValue(value, out int existingIndex))
            {
                return existingIndex;
            }

            // Otherwise, add new value
            SharedStringItem newItem = new SharedStringItem(new Text(value));
            shareStringPart.SharedStringTable.AppendChild(newItem);
            shareStringPart.SharedStringTable.Save();

            int newIndex = sharedStringLookup.Count;
            sharedStringLookup[value] = newIndex;

            return newIndex;
        }

        //Finds or creates a cell at the specified location.
        public static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);

            if (row == null)
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            string cellReference = columnName + rowIndex;

            Cell cell = row.Elements<Cell>()
                .FirstOrDefault(c => c.CellReference != null && c.CellReference.Value == cellReference);

            if (cell == null)
            {
                cell = new Cell() { CellReference = cellReference };
                // Insert in the correct position based on column
                Cell refCell = row.Elements<Cell>()
                    .FirstOrDefault(c => string.Compare(c.CellReference.Value, cellReference, true) > 0);

                row.InsertBefore(cell, refCell);
                worksheet.Save();
            }

            return cell;
        }

        public static WorksheetPart InsertWorksheet(WorkbookPart workbookPart,string worksheetName)
        {
            // Create a new worksheet part.  
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            // Get the Sheets collection from the workbook.  
            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            if (sheets == null)
            {
                sheets = workbookPart.Workbook.AppendChild(new Sheets());
            }

            // Generate a unique ID for the new sheet.  
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Any())
            {
                sheetId = sheets.Elements<Sheet>().Max(s => s.SheetId.Value) + 1;
            }

            // Create a new sheet and associate it with the worksheet part.  
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);
            Sheet sheet = new Sheet()
            {
                Id = relationshipId,
                SheetId = sheetId,
                Name = worksheetName,
            };
            sheets.Append(sheet);

            return newWorksheetPart;
        }
        public static int GetLastRowIndex(WorksheetPart worksheetPart)
        {
            if (worksheetPart.Worksheet.Descendants<Row>().Any())
            {
                return worksheetPart.Worksheet.Descendants<Row>().Max(r => (int)r.RowIndex.Value);
            }
            else
            {
                return 0; // No rows yet
            }
        }

        public static WorksheetPart GetWorksheetPartByName(WorkbookPart workbookPart, string sheetName)
        {
            foreach (var worksheetPart in workbookPart.WorksheetParts)
            {
                Sheet sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == sheetName);
                if (sheet != null)
                {
                    return worksheetPart;
                }
            }
            return null; // No worksheet found with the given name
        }
        // Helper method to get the first available worksheet
        public static WorksheetPart GetFirstWorksheetPart(WorkbookPart workbookPart)
        {
            return workbookPart.WorksheetParts.FirstOrDefault();
        }
        public static bool IsFileOpen(string filepath)
        {
            try
            {
                using (FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    // File is not open, we can access it.
                    return false;
                }
            }
            catch (IOException)
            {
                // File is open
                return true;
            }
        }
        public static string GetCellValue(Cell cell, SharedStringTablePart sharedStringPart)
        {
            if (cell == null || cell.CellValue == null)
                return "";

            string value = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (sharedStringPart != null)
                {
                    return sharedStringPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }

            return value; // For numbers, return as is
        }


    }
}
