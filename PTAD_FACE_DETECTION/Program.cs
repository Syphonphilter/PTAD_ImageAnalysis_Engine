using System;
using System.Collections.Generic;
using System.Drawing;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.PixelFormats;
using Microsoft.Azure.CognitiveServices.Vision.Face;
using Microsoft.Azure.CognitiveServices.Vision.Face.Models;
using OfficeOpenXml;
using Image = SixLabors.ImageSharp.Image;
using SixLabors.ImageSharp.Processing.Processors.Convolution;

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
            string enfolderPath = @"C:\BESPhotos\EnhancedImages\";

            // Create an Excel package
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage();
            

            // Set up the Face API client
            IFaceClient faceClient = new FaceClient(new ApiKeyServiceClientCredentials(SubscriptionKey)) { Endpoint = Endpoint };
            List<Tuple<string,string>> facesDetected = new List<Tuple<string, string>>();

            List<Tuple<string, string>> facesNotDetected = new List<Tuple<string, string>>();
            // Loop through each image file in the folder
            DirectoryInfo directory = new DirectoryInfo(folderPath);
            DirectoryInfo enhcanceddirectory = new DirectoryInfo(enfolderPath);
            foreach (FileInfo file in directory.GetFiles("*.jpg"))
            {
                string fileName = Path.GetFileNameWithoutExtension(file.Name);
                EnchanceImage(fileName,false);
            }
            foreach (FileInfo filec in enhcanceddirectory.GetFiles("*.jpg"))
            {
                // Load the image file
                using (Stream imageStream = File.OpenRead(filec.FullName))
                {
                    try
                    {
                        // Detect faces in the image
                        IList<DetectedFace> faces = faceClient.Face.DetectWithStreamAsync(
                            imageStream,
                            detectionModel: DetectionModel.Detection01,
                            recognitionModel: RecognitionModel.Recognition04,
                            returnFaceAttributes: new List<FaceAttributeType> { FaceAttributeType.QualityForRecognition, FaceAttributeType.Blur,FaceAttributeType.Exposure
                            }).Result;

                        // Add the file name and face detection result to the Excel sheet
                        string fileName = Path.GetFileNameWithoutExtension(filec.Name);
                      
                        if (faces.Count > 0)
                        {
                            double qualityScore = (1 - faces[0].FaceAttributes.Blur.Value) * (1 - faces[0].FaceAttributes.Exposure.Value);
                       
                            facesDetected.Add(Tuple.Create(fileName, qualityScore.ToString()));
                            if (qualityScore < 0.6)
                            {
                                EnchanceImage(fileName,true);
                            }

                        }
                        else
                        {
                           
                            facesNotDetected.Add(Tuple.Create(fileName, "0"));

                        }
                    }
                    catch(System.AggregateException e)
                    {
                        Console.WriteLine(e.InnerException);
                        Console.WriteLine(e.InnerExceptions);
                        Console.WriteLine(e.Data);
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
                foreach (var filename in facesDetected)
                {
                    worksheet.Cells[row, 1].Value = filename.Item1;
                    worksheet.Cells[row, 2].Value = "Yes";
                    worksheet.Cells[row, 3].Value = filename.Item2;
                    row++;
                }
                foreach (var filename in facesNotDetected)
                {
                    worksheet.Cells[row, 1].Value = filename;
                    worksheet.Cells[row, 2].Value = "No";
                    worksheet.Cells[row, 3].Value = filename.Item2;
                    row++;
                }

                // Save the Excel file
                FileInfo excelFile = new FileInfo(@"C:\Users\Syphonphilter\Documents\PTAD DOCS\FaceDetectionResults.xlsx");
                excel.SaveAs(excelFile);
            }
            // Save the Excel package
           // package.SaveAs(new FileInfo(@"C:\Users\Syphonphilter\Documents\PTAD DOCS\FaceDetectionResults.xlsx"));
        }


        public static async void EnchanceImage(string fileName, bool?secondphase)
        {
            // Specify the input and output folders
            string folderPath = "";
            if (secondphase== true)
            {
                 folderPath = @"C:\BESPhotos\EnhancedImages\";
            }
            else
            {
                 folderPath = @"C:\BESPhotos\TEST\";
            }
           

            string outputFolder = @"C:\BESPhotos\EnhancedImages\";
            string secondoutputFolder = @"C:\BESPhotos\ReEnhancedImages\";
           
            // Create the output folder if it doesn't exist
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            } if (!Directory.Exists(secondoutputFolder))
            {
                Directory.CreateDirectory(secondoutputFolder);
            }

            // Get the list of image files in the input folder




            try
            {
                // Load the image from the file path
                var  extension = ".jpg";
                
                using (Stream imageStream = File.OpenRead(folderPath + fileName+extension))
                {
                   using var image = Image.Load(folderPath + fileName + ".jpg");
                    // Detect faces in the image

                    if (secondphase == true)
                    {
                        image.Mutate(x => x.GaussianSharpen(1.1f));

                        // Apply a brightness filter to the image

                        image.Mutate(x => x.Brightness((float)(0.2) * 5));
                    }
                    else
                    {
                        image.Mutate(x => x.GaussianSharpen(1.5f));

                        // Apply a brightness filter to the image

                        image.Mutate(x => x.Brightness((float)(0.4) * 5));
                    }
                    string filePath = outputFolder+fileName + ".jpg";

                    if (secondphase == true)
                    {
                        string outputFilePath = Path.Combine(secondoutputFolder, Path.GetFileNameWithoutExtension(fileName) + ".jpg");
                        image.Save(outputFilePath);
                    }
                    else
                    {
                        string outputFilePath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(fileName) + ".jpg");
                        image.Save(outputFilePath);
                    }
                }

                // Save the enhanced image to the output folder
             

            }
            
            catch (Exception ex)
            {
                // Log any exceptions and continue to the next image
                Console.WriteLine($"Error processing image {fileName}: {ex.Message}");
                
            }
            }

        }
    }

