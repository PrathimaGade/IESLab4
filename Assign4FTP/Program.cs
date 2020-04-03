using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Assign4FTP.Models;
using Assign4FTP.Models.Utilities;
using System.Net;
using Newtonsoft.Json;
using Assign4FTP.Model;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using FTPAp.Models;
using Assign4FTP.Models.Utilities;
using System.Drawing.Imaging;

namespace Assign4FTP
{
    class Program
    {


        static void Main(string[] args)
        {
            Student myrecord = new Student { StudentId = "200430242", FirstName = "BalaPrathima", LastName = "Gade" };
            Student student1 = new Student();
            List<string> directories = FTP.GetDirectory(Constants.FTP.BaseUrl);
            List<Student> students = new List<Student>();
            foreach (var directory in directories)
            {
                Student student = new Student() { AbsoluteUrl = Constants.FTP.BaseUrl };
                student.UID = Guid.NewGuid().ToString();
                student.FromDirectory(directory);
                if (student.StudentId == "200430242")
                {
                    student.IsMe = true;
                }
                students.Add(student);
            }



            HttpWebRequest WebReq = (HttpWebRequest)WebRequest.Create(string.Format("https://jsonplaceholder.typicode.com/users"));

            WebReq.Method = "GET";

            HttpWebResponse WebResp = (HttpWebResponse)WebReq.GetResponse();

            Console.WriteLine(WebResp.StatusCode);
            Console.WriteLine(WebResp.Server);

            string jsonString;
            using (Stream stream = WebResp.GetResponseStream())   //modified from your code since the using statement disposes the stream automatically when done
            {
                StreamReader reader = new StreamReader(stream, System.Text.Encoding.UTF8);
                jsonString = reader.ReadToEnd();
            }
            List<Model.Users> users = JsonConvert.DeserializeObject<List<Model.Users>>(jsonString);

            Console.WriteLine(users.Count());
            Console.WriteLine(jsonString);

            /*-----------------------WORD DOCUMENT-------------------------------------- */

            string docxFilePath = $"{Constants.Locations.DataFolder}//info.docx";
            //string ftpImagePath = Constants.FTP.BaseUrl + "/200430242 BalaPrathima Gade/myimage.jpg";
            string studentsImagePath = $"{Constants.Locations.ImagesFolder}//myimage.jpg";
            // Create a document by supplying the filepath. 
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(docxFilePath, WordprocessingDocumentType.Document))
            {



                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());





                
                foreach (var user in users)
                {
                    run.AppendChild(new Text("Id:"));
                    run.AppendChild(new Text(user.Id.ToString()));
                    run.AppendChild(new Text("   "));
                    run.AppendChild(new Text("My Name is: "));
                    run.AppendChild(new Text(user.Name.ToString()));
                    run.AppendChild(new Text("    "));
                    run.AppendChild(new Text("My Email is: "));
                    run.AppendChild(new Text(user.Email.ToString()));
                    run.AppendChild(new Text("    "));
                    run.AppendChild(new Break());
                    using (FileStream stream = new FileStream(studentsImagePath, FileMode.Open))
                    {
                        imagePart.FeedData(stream);
                    }
                    Word.AddImageToBody(wordDocument, mainPart.GetIdOfPart(imagePart));



                    run.AppendChild(new Break() { Type = BreakValues.Page });
                }

            }
            /*-----------------------------EXCEL DOCUMENT-------------------------------------------*/

			string xlsxFilePath = $"{Constants.Locations.DataFolder}//info.xlsx";

			// Create a spreadsheet document by supplying the filepath.
			// By default, AutoSave = true, Editable = true, and Type = xlsx.
			SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(xlsxFilePath, SpreadsheetDocumentType.Workbook);

                                      //Creating Excel Document Structure//

			// Add a WorkbookPart to the document.
			WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
			workbookpart.Workbook = new Workbook();

			// Add a WorksheetPart to the WorkbookPart.
			WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
			worksheetPart.Worksheet = new Worksheet(new SheetData());

			// Add Sheets to the Workbook.
			Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                     
            SharedStringTablePart shareStringPart;
			shareStringPart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();

			WorksheetPart worksheetPart2 = workbookpart.AddNewPart<WorksheetPart>();
			worksheetPart2.Worksheet = new Worksheet(new SheetData());
			Sheet sheet = new Sheet()
			{
				Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart2),
				SheetId = 1,
				Name = "studentdata"
			};
			sheets.Append(sheet);

            //Creating Heading Row
            string[] headingRow = { "UID", "StudentID","FirstName", "LastName", "DateofBirth", "IsMe", "Age" };
            char headingIndex = 'A';
            for(int i=0;i< headingRow.Count(); i++)
            {
                Excel.InsertCellInWorksheet(headingIndex.ToString(), 1, worksheetPart2).CellValue = new CellValue(Excel.InsertSharedStringItem(headingRow[i], shareStringPart).ToString());
                Excel.InsertCellInWorksheet(headingIndex.ToString(), 1, worksheetPart2).DataType = new EnumValue<CellValues>(CellValues.SharedString);
                headingIndex++;
            }

            //Processing student Data 
            uint rowIndex = 2;
            foreach (var student in students)
            {
                char columnIndex = 'A';
                int uidIndex = Excel.InsertSharedStringItem(student.UID.ToString(), shareStringPart);
                Cell uidCell = Excel.InsertCellInWorksheet(columnIndex.ToString(), rowIndex, worksheetPart2);
                uidCell.CellValue = new CellValue(uidIndex.ToString());
                uidCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                columnIndex++;



                string ID = student.StudentId;

                int studentIdIndex = Excel.InsertSharedStringItem(ID, shareStringPart);
                Cell studentCell = Excel.InsertCellInWorksheet(columnIndex.ToString(), rowIndex, worksheetPart2);
                studentCell.CellValue = new CellValue(studentIdIndex.ToString());
                studentCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                columnIndex++;


                int firstIndex = Excel.InsertSharedStringItem(student.FirstName.ToString(), shareStringPart);
                Cell firstCell = Excel.InsertCellInWorksheet(columnIndex.ToString(), rowIndex, worksheetPart2);
                firstCell.CellValue = new CellValue(firstIndex.ToString());
                firstCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                columnIndex++;



                int Lastindex = Excel.InsertSharedStringItem(student.LastName.ToString(), shareStringPart);
                Cell LastCell = Excel.InsertCellInWorksheet(columnIndex.ToString(), rowIndex, worksheetPart2);
                LastCell.CellValue = new CellValue(Lastindex.ToString());
                LastCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                columnIndex++;


                int DOBIndex = Excel.InsertSharedStringItem(student.DateOfBirthDT.ToShortDateString(), shareStringPart);
                Cell DOBCell = Excel.InsertCellInWorksheet(columnIndex.ToString(), rowIndex, worksheetPart2);
                DOBCell.CellValue = new CellValue(DOBIndex.ToString());
                DOBCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                columnIndex++;


                int IsMeIndex = Excel.InsertSharedStringItem(student.IsMe.ToString(), shareStringPart);
                Cell IsMeCell = Excel.InsertCellInWorksheet(columnIndex.ToString(), rowIndex, worksheetPart2);
                IsMeCell.CellValue = new CellValue(IsMeIndex.ToString());
                IsMeCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                columnIndex++;


                int ageIndex = Excel.InsertSharedStringItem(student.Age.ToString(), shareStringPart);
                Cell ageCell = Excel.InsertCellInWorksheet(columnIndex.ToString(), rowIndex, worksheetPart2);
                ageCell.CellValue = new CellValue(ageIndex.ToString());
                ageCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                rowIndex++;
            }


            // Save and Close the document.
            workbookpart.Workbook.Save();
			spreadsheetDocument.Close();
		}





        ///// this inserts a new worksheet, need to find a way to have it edit existing. Need one function for create a new sheet, and one for edit existing.
        //public static void InsertText(string docName, string text, uint rownum, string colletter)
        //{
        //    // Open the document for editing.
        //    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
        //    {
        //        // Get the SharedStringTablePart. If it does not exist, create a new one.
        //        SharedStringTablePart shareStringPart;
        //        if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
        //        {
        //            shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
        //        }
        //        else
        //        {
        //            shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
        //        }


        //        // Insert the text into the SharedStringTablePart.
        //        int index = InsertSharedStringItem(text, shareStringPart);


        //        // Insert a new worksheet.
        //        WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);


        //        // Insert cell A1 into the new worksheet.
        //        Cell cell = InsertCellInWorksheet(colletter, rownum, worksheetPart);


        //        // Set the value of cell A1.
        //        cell.CellValue = new CellValue(index.ToString());
        //        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);


        //        // Save the new worksheet.
        //        worksheetPart.Worksheet.Save();
        //        spreadSheet.Close();

        //    }
        //}

    }

    
}