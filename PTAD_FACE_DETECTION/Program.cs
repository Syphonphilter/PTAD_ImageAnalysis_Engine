using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Azure.CognitiveServices.Vision.Face;
using Microsoft.Azure.CognitiveServices.Vision.Face.Models;
using OfficeOpenXml;

namespace FaceDetection
{
    class Program
    {
        // Replace with your own subscription key and endpoint
        private const string SubscriptionKey = "a33340fd28644f3aa66c5b934abc094a";
        private const string Endpoint = "https://ptad.cognitiveservices.azure.com/";

        static void Main(string[] args)
        {
            // Replace with the path to your folder of headshot images
            string folderPath = @"C:\BESPhotos\TEST\";

            // Create an Excel package
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage();
            

            // Set up the Face API client
            IFaceClient faceClient = new FaceClient(new ApiKeyServiceClientCredentials(SubscriptionKey)) { Endpoint = Endpoint };
            List<string> facesDetected = new List<string>();
            List<string> facesNotDetected = new List<string>();
            // Loop through each image file in the folder
            DirectoryInfo directory = new DirectoryInfo(folderPath);
            foreach (FileInfo file in directory.GetFiles("*.jpg"))
            {
                // Load the image file
                using (Stream imageStream = File.OpenRead(file.FullName))
                {
                    // Detect faces in the image
                    IList<DetectedFace> faces = faceClient.Face.DetectWithStreamAsync(imageStream, detectionModel: DetectionModel.Detection03).Result;

                    // Add the file name and face detection result to the Excel sheet
                    string fileName = Path.GetFileNameWithoutExtension(file.Name);
                    if (faces.Count > 0)
                    {
                        facesDetected.Add(fileName);
                       
                    }
                    else
                    {
                        facesNotDetected.Add(fileName);
                       
                    }
                }
            }
            using (ExcelPackage excel = new ExcelPackage())
            {
                // Create a new worksheet
                ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("Headshot Report");

                // Add headers to the worksheet
                worksheet.Cells[1, 1].Value = "PensionerID";
                worksheet.Cells[1, 2].Value = "Face Detected";

                // Add data to the worksheet
                int row = 2;
                foreach (string filename in facesDetected)
                {
                    worksheet.Cells[row, 1].Value = filename;
                    worksheet.Cells[row, 2].Value = "Yes";
                    row++;
                }
                foreach (string filename in facesNotDetected)
                {
                    worksheet.Cells[row, 1].Value = filename;
                    worksheet.Cells[row, 2].Value = "No";
                    row++;
                }

                // Save the Excel file
                FileInfo excelFile = new FileInfo(@"C:\Users\Syphonphilter\Documents\PTAD DOCS\FaceDetectionResults.xlsx");
                excel.SaveAs(excelFile);
            }
            // Save the Excel package
           // package.SaveAs(new FileInfo(@"C:\Users\Syphonphilter\Documents\PTAD DOCS\FaceDetectionResults.xlsx"));
        }
    }
}
