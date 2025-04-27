using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadSheet;

namespace SpreadSheetTest
{
    [TestFixture]
    public class CreateTest
    {
        private string _tempFilePath;

        [SetUp]
        public void SetUp()
        {
            string folderPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string assemblyPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
            string[] assemblyPathSplit = assemblyPath.Split('\\');
            string sourceCodePath = "";

            for (int i = 1; i < assemblyPathSplit.Length - 3; i++)
            {
                if (i == 1)
                {
                    sourceCodePath = sourceCodePath + assemblyPathSplit[i];
                }
                else
                {
                    sourceCodePath = sourceCodePath + '\\' + assemblyPathSplit[i];
                }
            }
            string dlxResFolderPath = Path.Combine(sourceCodePath, "UnitTests");
            _tempFilePath = Path.Combine(sourceCodePath, Path.GetRandomFileName() + ".xlsx");
        }

        [Test]
        public void CreateSpreadsheetWorkbook_ShouldCreateFile()
        {
            // Act
            Create.CreateSpreadsheetWorkbook(_tempFilePath);

            // Assert
            Assert.IsTrue(File.Exists(_tempFilePath), "File should exist after creation.");

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(_tempFilePath, false))
            {
                Assert.NotNull(document.WorkbookPart);
                Assert.NotNull(document.WorkbookPart.Workbook);
                Assert.That(document.WorkbookPart.Workbook.Sheets.Elements<Sheet>().Any(), Is.True, "Workbook should contain at least one sheet.");
            }
        }
        [Test]
        public void InsertData_ShouldInsertDataIntoSheet()
        {
            // Arrange
            List<List<string>> inputData = new List<List<string>>
            {
                new List<string> { "Name", "Age" },
                new List<string> { "Alice", "30" },
                new List<string> { "Bob", "25" }
            };
            string sheetName = "TestSheet";

            // Act
            Create.InsertData(_tempFilePath, inputData, sheetName);

            // Assert
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(_tempFilePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
                Assert.IsNotNull(sheet, "Sheet should exist.");

                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                var rows = sheetData.Elements<Row>().ToList();
                Assert.That(rows.Count, Is.EqualTo(3), "There should be 3 rows (1 header + 2 data rows).");

                var firstCell = rows[1].Elements<Cell>().First();
                Assert.IsNotNull(firstCell.CellValue);
            }
        }
        [Test]
        public void ReadExcelFileDOM_ShouldPrintExcelContent()
        {
            // Arrange
            List<List<string>> inputData = new List<List<string>>
            {
                new List<string> { "Col1", "Col2" },
                new List<string> { "Data1", "Data2" }
            };
            Create.InsertData(_tempFilePath, inputData, "Sheet1");

            using (var sw = new StringWriter())
            {
                Console.SetOut(sw);

                // Act
                Create.ReadExcelFileDOM(_tempFilePath);

                // Assert
                var output = sw.ToString();
                Assert.IsTrue(output.Contains("Data1"), "Should contain inserted data.");
                Assert.IsTrue(output.Contains("Data2"));
            }
        }


        [TearDown]
        public void TearDown()
        {
            if (File.Exists(_tempFilePath))
                File.Delete(_tempFilePath);
        }
    }
}