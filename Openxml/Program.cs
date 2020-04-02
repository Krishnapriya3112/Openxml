using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using Openxml.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;

namespace Openxml
{
    class Program
    {
        static string MyData;
        static List<Student> myList;

        static void Main(string[] args)
        {
            MyData = FTP.GetFile();
            myList = MyApiData();
            ToDoc();
            excel.CreateSpreadsheet();
          
            FTP.UploadFile(@"D:\info.docx", "info.docx");

            
            FTP.UploadFile(@"D:\info.xlsx", "info.xlsx");
            
            Console.ReadLine();
        }
        public static string getMyData()
        {
            return MyData;
        }
        public static List<Student> MyApiData()
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://webapibasicsstudenttracker.azurewebsites.net/api/students");
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream());
            string json = reader.ReadToEnd();
            List<Student> studentList = JsonConvert.DeserializeObject<List<Student>>(json);
            return studentList;

        }
        static void ToDoc()
        {
            string data = getMyData();
            string[] record = data.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            string[] myrecord = record[1].Split(',');


            List<Student> studentList = myList;
            using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Create(@"D:\info.docx", WordprocessingDocumentType.Document))
            {

                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
               
                for (int i = 0; i < 10; i++)
                {
                    Run run = new Run();
                    if (i == 0)
                    {
                        run.AppendChild(new Text("Student ID:" + myrecord[0]));
                        run.AppendChild(new Break());
                        run.AppendChild(new Text("First Name:" + myrecord[1]));
                        run.AppendChild(new Break());
                        run.AppendChild(new Text("Last Name:" + myrecord[2]));
                        run.AppendChild(new Break());
                       
                    }
                    else
                    {
                        run.AppendChild(new Text("Student Code:" + studentList[i].StudentCode));
                        run.AppendChild(new Break());
                        run.AppendChild(new Text("First Name:" + studentList[i].FirstName));
                        run.AppendChild(new Break());
                        run.AppendChild(new Text("Last Name:" + studentList[i].LastName));
                    }
                    if (i < 9)
                        run.AppendChild(new Break() { Type = BreakValues.Page });
                    Paragraph para = new Paragraph(run);
                    mainPart.Document.Body.InsertAfter(para, mainPart.Document.Body.LastChild);

                }
                mainPart.Document.Save();
            }

            Console.WriteLine("Doc one...");

        }
    }
}
