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

namespace Assign4FTP
{
    class Program
    {


        static void Main(string[] args)
        {
            Student myrecord = new Student { StudentId = "200430242", FirstName = "BalaPrathima", LastName = "Gade" };

            List<string> directories = FTP.GetDirectory(Constants.FTP.BaseUrl);
            List<Student> students = new List<Student>();

            foreach (var directory in directories)
            {
                Student student = new Student() { AbsoluteUrl = Constants.FTP.BaseUrl };
                student.FromDirectory(directory);
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
            List<Users> items = JsonConvert.DeserializeObject<List<Users>>(jsonString);

            Console.WriteLine(items.Count());
            Console.WriteLine(jsonString);


            string docxFilePath = $"{Constants.Locations.DataFolder}//info.docx";
            string ftpImagePath = $"/" ;
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
                //using (FileStream stream = new FileStream(studentsImagePath, FileMode.Open))
                //{
                //    imagePart.FeedData(stream);
                //}

                //Word.AddImageToBody(wordDocument, mainPart.GetIdOfPart(imagePart));
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
               

                

                

                foreach (var student in items)
                {
                    run.AppendChild(new Text("Id:"));
                    run.AppendChild(new Text(student.Id.ToString()));
                    run.AppendChild(new Text("    "));
                    run.AppendChild(new Text("My Name is: "));
                    run.AppendChild(new Text(student.Name.ToString()));
                    run.AppendChild(new Text("    "));
                    run.AppendChild(new Break() { Type = BreakValues.Page });


                    //if (student.Record == true)
                    //{
                        using (FileStream stream = new FileStream(studentsImagePath, FileMode.Open))
                        {
                            imagePart.FeedData(stream);
                        }
                        AddImageToBody(wordDocument, mainPart.GetIdOfPart(imagePart));
                    //}


                }

            }
           // using (WordprocessingDocument wordprocessingDocument =
           //WordprocessingDocument.Open(docxFilePath, true))
           // {
           //     MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

           //     ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

           //     using (FileStream stream = new FileStream(studentsImagePath, FileMode.Open))
           //     {
           //         imagePart.FeedData(stream);
           //     }

           //     AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
           // }







            //string studentsjsonPath = $"{Constants.Locations.DataFolder}//students.json";
            ////Establish a file stream to collect data from the response
            //using (StreamWriter fs = new StreamWriter(studentsjsonPath))
            //{
            //    foreach (var student in students)
            //    {
            //        string Student = Newtonsoft.Json.JsonConvert.SerializeObject(student);
            //        fs.WriteLine(Student.ToString());
            //        //Console.WriteLine(jStudent);
            //    }
            //}



            //string studentsExcelPath = $"{Constants.Locations.DataFolder}//students.xlsx";

            //SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
            //Create(studentsExcelPath, SpreadsheetDocumentType.Workbook);

            //// Add a WorkbookPart to the document.
            //WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            //workbookpart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

            //// Add a WorksheetPart to the WorkbookPart.
            //WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            //worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new DocumentFormat.OpenXml.Spreadsheet.SheetData());

            //// Add Sheets to the Workbook.
            //DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
            //    AppendChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>(new DocumentFormat.OpenXml.Spreadsheet.Sheets());

            //// Append a new worksheet and associate it with the workbook.
            //DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet()
            //{
            //    Id = spreadsheetDocument.WorkbookPart.
            //    GetIdOfPart(worksheetPart),
            //    SheetId = 1,
            //    Name = "mySheet"
            //};
            //sheets.Append(sheet);

            //workbookpart.Workbook.Save();

            //// Close the document.
            //spreadsheetDocument.Close();



            //    string studentsxmlPath = $"{Constants.Locations.DataFolder}//students.xml";
            //    //Establish a file stream to collect data from the response
            //    using (StreamWriter fs = new StreamWriter(studentsxmlPath))
            //    {
            //        XmlSerializer x = new XmlSerializer(students.GetType());
            //        x.Serialize(fs, students);
            //        Console.WriteLine();
            //    }

           
           

            //    return;


        }


        //        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

        //        using (FileStream stream = new FileStream(fileName, FileMode.Open))
        //        {
        //            imagePart.FeedData(stream);
        //        }

        //        AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
        //    }
        //}

        public static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            // Define the reference of the image.
            var element =
                 new DocumentFormat.OpenXml.Wordprocessing.Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = 990000L, Cy = 792000L },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run(element)));
        }


        //public static Image Base64ToImage(string base64String)
        //{
        //    // Convert Base64 String to byte[]
        //    byte[] imageBytes = Convert.FromBase64String(base64String.Trim());
        //    var ms = new MemoryStream(imageBytes, 0, imageBytes.Length);
        //    // Convert byte[] to Image
        //    ms.Write(imageBytes, 0, imageBytes.Length);
        //    System.Drawing.Image image = System.Drawing.Image.FromStream(ms, true);
        //    return image;
        //}

    }
}