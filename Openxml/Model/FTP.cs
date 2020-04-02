using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

namespace Openxml.Model
{
    class FTP
    {
        const string UserName = @"bdat100119f\bdat1001";
        const string Password = "bdat1001";

        const string URL = "ftp://waws-prod-dm1-127.ftp.azurewebsites.windows.net/bdat1001-20914/200450333%20Krishnapriya%20Sarojam/info.csv";
        const string URL1 = "ftp://waws-prod-dm1-127.ftp.azurewebsites.windows.net/bdat1001-20914/";
        const string ImageURL = "ftp://waws-prod-dm1-127.ftp.azurewebsites.windows.net/bdat1001-20914/200450333%20Krishnapriya%20Sarojam/myimage.jpg";
        public static byte[] GetImage(string url = ImageURL, string username = UserName, string password = Password)
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(url);
            request.Credentials = new NetworkCredential(username, password);
            request.Method = WebRequestMethods.Ftp.DownloadFile;
            request.EnableSsl = false;
            FtpWebResponse response = (FtpWebResponse)request.GetResponse();
            Stream responseStream = response.GetResponseStream();
            var data = ToByteArray(responseStream);
            Console.WriteLine($"Download Complete, status {response.StatusDescription}");

            return data;
        }
        public static byte[] ToByteArray(Stream stream)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                byte[] chunk = new byte[1024];
                int bytesRead;
                while ((bytesRead = stream.Read(chunk, 0, chunk.Length)) > 0)
                {
                    ms.Write(chunk, 0, bytesRead);
                }

                return ms.ToArray();
            }
        }
        public static string GetFile(string url = URL, string username = UserName, string password = Password)
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(url);
            request.Credentials = new NetworkCredential(username, password);
            request.Method = WebRequestMethods.Ftp.DownloadFile;
            request.EnableSsl = false;
            FtpWebResponse response = (FtpWebResponse)request.GetResponse();

            Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream);
            string data = reader.ReadToEnd();

            Console.WriteLine($"Download Complete, status {response.StatusDescription}");

            reader.Close();
            response.Close();
            return data;
        }
        public static string GetStudentFolder(string url = URL1, string username = UserName, string password = Password)
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(url);
            request.Credentials = new NetworkCredential(username, password);
            request.Method = WebRequestMethods.Ftp.ListDirectory;
            request.EnableSsl = false;
            FtpWebResponse response = (FtpWebResponse)request.GetResponse();

            Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream);
            string data = reader.ReadToEnd();

            Console.WriteLine($"Download Complete, status {response.StatusDescription}");

            reader.Close();
            response.Close();

            return data;
        }
        public static string GetDataFromStudentFolder(string url, string username = UserName, string password = Password)
        {
            try

            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(URL1 + url);
                request.Credentials = new NetworkCredential(username, password);
                request.Method = WebRequestMethods.Ftp.DownloadFile;
                request.EnableSsl = false;
                FtpWebResponse response = (FtpWebResponse)request.GetResponse();

                Stream responseStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(responseStream);
                string data = reader.ReadToEnd();

                Console.WriteLine($"Download Complete, status {response.StatusDescription}");

                reader.Close();
                response.Close();
                //Console.WriteLine( data);
                return data;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                
            }
            return null;
        }
        public static string UploadFile(string sourceFilePath, string destinationFile, string username = UserName, string password = Password)
        {
            string output;
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://waws-prod-dm1-127.ftp.azurewebsites.windows.net/bdat1001-20914/200450333%20Krishnapriya%20Sarojam/" + destinationFile);
            request.Method = WebRequestMethods.Ftp.UploadFile;
            request.Credentials = new NetworkCredential(username, password);

            byte[] fileContents = Encoding.UTF8.GetBytes(sourceFilePath);

            //Get the length or size of the file
            request.ContentLength = fileContents.Length;

            //Write the file to the stream on the server
            using (Stream requestStream = request.GetRequestStream())
            {
                requestStream.Write(fileContents, 0, fileContents.Length);
            }

            //Send the request
            using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
            {
                output = $"Upload File Complete, status {response.StatusDescription}";
            }
            //Thread.Sleep(Constants.FTP.OperationPauseTime);

            return (output);
        }
    }
}
