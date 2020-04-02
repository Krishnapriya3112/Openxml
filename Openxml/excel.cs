using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Openxml.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Openxml
{
    class excel
    {
        public static void CreateSpreadsheet()
        {

            SpreadsheetDocument spreadsheetDocument =
            SpreadsheetDocument.Create(@"D:\info.xlsx", SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

          
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

          
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sheet1"
            };
            
            sheets.Append(sheet);
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            SharedStringTablePart shareStringPart;
            if (spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }
            int index = InsertSharedStringItem("Hello,my name is Krishnapriya Sarojam", shareStringPart);

            Cell cell = InsertCellInWorksheet("A", 2, worksheetPart);
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            worksheetPart = InsertWorksheet(workbookpart);
          
            string data = FTP.GetStudentFolder();
            int flag = 0;
            if (data != null)
            {
                string[] studentFolder = data.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                uint count = 65;
                uint row = 2;
                foreach (string f in studentFolder)
                {
                    if (FTP.GetDataFromStudentFolder(f + "\\info.csv") != null)
                    {
                        string[] record = FTP.GetDataFromStudentFolder(f + "\\info.csv").Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                        string[] recordHeading = record[0].Split(',');
                        string[] recordData = record[1].Split(',');
                        var result = char.ConvertFromUtf32((int)count);
                        if (flag == 0)
                        {
                            InsertCell(shareStringPart, worksheetPart, result, 1, index, cell, "UniqueData", CellValues.String);
                            count++;
                            foreach (string s in recordHeading)
                            {
                                if (s != "ImageData")
                                {
                                    result = char.ConvertFromUtf32((int)count);
                                    InsertCell(shareStringPart, worksheetPart, result, 1, index, cell, s, CellValues.SharedString);
                                    count++;
                                }
                            }
                            flag = 1;
                            result = char.ConvertFromUtf32((int)count);
                            InsertCell(shareStringPart, worksheetPart, result, 1, index, cell, "Age", CellValues.SharedString);
                            count++;
                            result = char.ConvertFromUtf32((int)count);
                            InsertCell(shareStringPart, worksheetPart, result, 1, index, cell, "IsMe", CellValues.SharedString);
                            count++;

                        }
                        count = 65;
                        result = char.ConvertFromUtf32((int)count);
                        InsertCell(shareStringPart, worksheetPart, result, row, index, cell, Guid.NewGuid().ToString(), CellValues.SharedString);
                        count++;
                        for (int i = 0; i <= 5; i++)
                        {
                            if (i < 3)
                            {
                                result = char.ConvertFromUtf32((int)count);
                                InsertCell(shareStringPart, worksheetPart, result, row, index, cell, recordData[i], CellValues.SharedString);
                            }
                            else if (i == 3)
                            {
                                try

                                {
                                    result = char.ConvertFromUtf32((int)count);
                                    InsertCell(shareStringPart, worksheetPart, result, row, index, cell, recordData[i], CellValues.SharedString);
                                }
                                catch
                                {
                                    string[] datev = recordData[i].ToString().Split(new string[] { "/", "-" }, StringSplitOptions.None);
                                    string temp = datev[0];
                                    datev[0] = datev[1];
                                    datev[1] = temp;
                                    DateTime date = new DateTime(int.Parse(datev[2]), int.Parse(datev[1]), int.Parse(datev[0]));
                                    Console.WriteLine(date.ToString("MM-dd-yyyy"));
                                    result = char.ConvertFromUtf32((int)count);
                                
                                    index = InsertSharedStringItem(date.ToString("MM-dd-yyyy"), shareStringPart);

                                    cell = InsertCellInWorksheet(result, row, worksheetPart);

                                    
                                    cell.CellValue = new CellValue(date.ToString("dd-MM-yyyy"));
                                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                                }

                            }
                            else if (i == 4)
                            {

                                // index = InsertSharedStringItem("=(TODAY()-E2) / 365", shareStringPart);
                                result = char.ConvertFromUtf32((int)count);
                                Console.WriteLine("Result:" + result);
                                cell = InsertCellInWorksheet(result, row, worksheetPart);
                                //  cell.CellValue = new CellValue(index.ToString());
                                cell.CellFormula = new CellFormula();
                                Console.WriteLine("Formula:" + char.ConvertFromUtf32((char.Parse(result) - 1)));
                                cell.CellFormula.Text = "(TODAY() - " + (char.ConvertFromUtf32((char.Parse(result) - 1))) + row.ToString() + ") / 365";
                                Console.WriteLine(cell.CellFormula.Text);
                                //  cell.
                                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                            }
                            else
                            {
                                result = char.ConvertFromUtf32((int)count);
                                index = InsertSharedStringItem("0", shareStringPart);

                                cell = InsertCellInWorksheet(result, row, worksheetPart);

                                //Console.WriteLine(recordData[i].ToString());
                                if (recordData[0] == "200450333")
                                    cell.CellValue = new CellValue("1");
                                else
                                    cell.CellValue = new CellValue("0");
                                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            }

                            count++;

                        }

                        row++;
                    }
                }
            }
            //end here

            workbookpart.Workbook.Save();

            
            spreadsheetDocument.Close();
            Console.WriteLine("Excel Done");
           

        }
        public static void InsertCell(SharedStringTablePart shareStringPart, WorksheetPart worksheetPart, string result, uint rowindex, int index, Cell cell, string data, CellValues cellValue)
        {
            index = InsertSharedStringItem(data, shareStringPart);
            cell = InsertCellInWorksheet(result, rowindex, worksheetPart);
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(cellValue);
        }
        
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();

            }

            int i = 0;

           
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }
        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetName = "Sheet" + sheetId;

           
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
               
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }
    }
}
