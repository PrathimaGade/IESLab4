﻿using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Assign4FTP.Models
{
    public class Excel
    {
		public static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
		{
			// If the part does not contain a SharedStringTable, create one.
			if (shareStringPart.SharedStringTable == null)
			{
				shareStringPart.SharedStringTable = new SharedStringTable();
			}


			int i = 0;


			// Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
			foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
			{
				if (item.InnerText == text)
				{
					return i;
				}


				i++;
			}


			// The text does not exist in the part. Create the SharedStringItem and return its index.
			shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
			shareStringPart.SharedStringTable.Save();


			return i;
		}


		public static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
		{
			Worksheet worksheet = worksheetPart.Worksheet;
			SheetData sheetData = worksheet.GetFirstChild<SheetData>();
			string cellReference = columnName + rowIndex;


			// If the worksheet does not contain a row with the specified row index, insert one.
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


			// If there is not a cell with the specified column name, insert one.  
			if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
			{
				return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
			}
			else
			{
				// Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
				Cell refCell = null;
				foreach (Cell cell in row.Elements<Cell>())
				{
					if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
					{
						refCell = cell;
						break;
					}
				}


				Cell newCell = new Cell() { CellReference = cellReference };
				row.InsertBefore(newCell, refCell);


				worksheet.Save();
				return newCell;
			}
		}


		// Given a WorkbookPart, inserts a new worksheet.


		public static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
		{
			// Add a new worksheet part to the workbook.
			WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
			newWorksheetPart.Worksheet = new Worksheet(new SheetData());
			newWorksheetPart.Worksheet.Save();


			Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
			string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);


			// Get a unique ID for the new sheet.
			uint sheetId = 1;
			if (sheets.Elements<Sheet>().Count() > 0)
			{
				sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
			}


			string sheetName = "Sheet" + sheetId;


			// Append the new worksheet and associate it with the workbook.
			Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
			sheets.Append(sheet);
			workbookPart.Workbook.Save();


			return newWorksheetPart;
		}




		/// this inserts a new worksheet, need to find a way to have it edit existing. Need one function for create a new sheet, and one for edit existing.
		public static void InsertText(string docName, string text, uint rownum, string colletter)
		{
			// Open the document for editing.
			using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
			{
				// Get the SharedStringTablePart. If it does not exist, create a new one.
				SharedStringTablePart shareStringPart;
				if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
				{
					shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
				}
				else
				{
					shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
				}


				// Insert the text into the SharedStringTablePart.
				int index = InsertSharedStringItem(text, shareStringPart);


				// Insert a new worksheet.
				WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);


				// Insert cell A1 into the new worksheet.
				Cell cell = InsertCellInWorksheet(colletter, rownum, worksheetPart);


				// Set the value of cell A1.
				cell.CellValue = new CellValue(index.ToString());
				cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);


				// Save the new worksheet.
				worksheetPart.Worksheet.Save();
				spreadSheet.Close();


			}
		}
	}
}
