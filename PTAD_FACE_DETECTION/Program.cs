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
            string folderPath = @"C:\PPDPhotos\";
            string enfolderPath = @"C:\PPDENHANCEDIMAGES\";

            // Create an Excel package
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage();
            

            // Set up the Face API client
            IFaceClient faceClient = new FaceClient(new ApiKeyServiceClientCredentials(SubscriptionKey)) { Endpoint = Endpoint };
            List<Tuple<string,string,string,string,string,string,string>> facesDetected = new List<Tuple<string, string, string, string, string,string,string>>();

            List<Tuple<string, string, string, string, string, string, string>> facesNotDetected = new List<Tuple<string, string, string, string, string, string, string>>();
            // Loop through each image file in the folder
            DirectoryInfo directory = new DirectoryInfo(folderPath);
            DirectoryInfo enhcanceddirectory = new DirectoryInfo(enfolderPath);
            foreach (FileInfo file in directory.GetFiles("*.jpg"))
            {
                string fileName = Path.GetFileNameWithoutExtension(file.Name);
                EnchanceImage(fileName, false);
            }
            foreach (FileInfo filec in directory.GetFiles("*.jpg"))
            {
                // Load the image file
                using (Stream imageStream = File.OpenRead(filec.FullName))
                {
                    using (Stream stream = File.OpenRead(filec.FullName))
                    {
                        try
                        {
                            // Detect faces in the image
                            IList<DetectedFace> faces = faceClient.Face.DetectWithStreamAsync(
                                imageStream,
                                detectionModel: DetectionModel.Detection03,
                                recognitionModel: RecognitionModel.Recognition04,
                               returnFaceAttributes: new List<FaceAttributeType> { FaceAttributeType.QualityForRecognition }
                               ).Result;
                            IList<DetectedFace> faces1 = faceClient.Face.DetectWithStreamAsync(
                              stream,
                              detectionModel: DetectionModel.Detection01,
                              recognitionModel: RecognitionModel.Recognition04,
                             returnFaceAttributes: new List<FaceAttributeType> { FaceAttributeType.QualityForRecognition, FaceAttributeType.Blur, FaceAttributeType.Exposure, FaceAttributeType.Noise, FaceAttributeType.Occlusion, FaceAttributeType.QualityForRecognition, FaceAttributeType.Accessories }
                             ).Result;


                            foreach (var faceDetection03 in faces)
                            {
                                var matchingFaceDetection01 = faces1.FirstOrDefault(x => x.FaceId == faceDetection03.FaceId);
                                if (matchingFaceDetection01 != null)
                                {
                                    // Combine the face attributes from the two detection models
                                    faceDetection03.FaceAttributes = matchingFaceDetection01.FaceAttributes;
                                }
                            }

                            // Use the combined list of faces with attributes
                            IList<DetectedFace> finalfaces = faces;
                            // Add the file name and face detection result to the Excel sheet
                            string fileName = Path.GetFileNameWithoutExtension(filec.Name);

                            if (faces.Count > 0)
                            {
                                double qualityScore = 0;
                                string noise = "";
                                string Eyeocculsion = "";
                                string Quality = "";
                                string Mouthocculsion = "";
                                string ForeHeadocculsion = "";

                               
                                if ( faces[0].FaceAttributes.Exposure== null)
                                {

                                     qualityScore =0;
                                     noise = 0.ToString();
                                     Eyeocculsion = 0.ToString();
                                    if(faces[0].FaceAttributes.QualityForRecognition == null)
                                    {
                                        Quality = "Low";
                                    }
                                    else
                                    {
                                        Quality = "Below Medium";
                                        EnchanceImage(fileName, true);
                                    }
                                  
                                    Mouthocculsion = 0.ToString();
                                    ForeHeadocculsion = 0.ToString();
                                }
                                else
                                {
                                    var blur = faces[0].FaceAttributes.Blur.Value;
                                    var exposure = faces[0].FaceAttributes.Exposure.Value;
                                    var noiseval = faces[0].FaceAttributes.Noise.Value;
                                    var occlusion = faces[0].FaceAttributes.Occlusion;
                                    var occlusionWeight = 0.1;
                                    var weightedSum = (blur * 0.2) + (exposure * 0.3) + (noiseval * 0.2);

                                    if (occlusion.ForeheadOccluded)
                                    {
                                        weightedSum += occlusionWeight;
                                    }

                                    if (occlusion.EyeOccluded)
                                    {
                                        weightedSum += occlusionWeight;
                                    }

                                    if (occlusion.MouthOccluded)
                                    {
                                        weightedSum += occlusionWeight;
                                    }
                                    qualityScore = (10 - weightedSum) / 10;
                                    var qualityForRecognition = faces[0].FaceAttributes.QualityForRecognition.ToString();
                                    var harmonizedScore = 0.0;

                                    if (qualityForRecognition == "Low")
                                    {
                                        harmonizedScore = 1.0;
                                    }
                                    else if (qualityForRecognition == "Medium")
                                    {
                                        harmonizedScore = 5.0;
                                    }
                                    else if (qualityForRecognition == "High")
                                    {
                                        harmonizedScore = 10.0;
                                    }

                                    var harmonizedQualityScore = (qualityScore + harmonizedScore) / 2.0;


                                    qualityScore = harmonizedQualityScore;

                                    noise = faces[0].FaceAttributes.Noise.Value.ToString();
                                     Eyeocculsion = faces[0].FaceAttributes.Occlusion.EyeOccluded.ToString();
                                     Quality = faces[0].FaceAttributes.QualityForRecognition.ToString();
                                     Mouthocculsion = faces[0].FaceAttributes.Occlusion.MouthOccluded.ToString();
                                     ForeHeadocculsion = faces[0].FaceAttributes.Occlusion.ForeheadOccluded.ToString();
                                }


                                facesDetected.Add(Tuple.Create(fileName, qualityScore.ToString(), noise, Eyeocculsion, Mouthocculsion, ForeHeadocculsion, Quality));
                             

                            }
                            else
                            {


                                facesNotDetected.Add(Tuple.Create(fileName, "0", "0", "0", "0", "0", "0"));

                            }
                        }
                        catch (System.AggregateException e)
                        {
                            Console.WriteLine(e.InnerException);
                            Console.WriteLine(e.InnerExceptions);
                            Console.WriteLine(e.Data);
                        }
                    }
                }
            }
            using (ExcelPackage excel = new ExcelPackage())
            {
                // Create a new worksheet
                ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("Headshot Report");

                // Add headers to the worksheet
                worksheet.Cells[1, 1].Value = "PensionerID";
                worksheet.Cells[1, 2].Value = "Face Detection Result";
                worksheet.Cells[1, 3].Value = "Quality Score";
                worksheet.Cells[1, 4].Value = "Noise";
                worksheet.Cells[1, 5].Value = "EyeOcculsion";
                worksheet.Cells[1, 6].Value = "MouthOcculsion";
                worksheet.Cells[1, 7].Value = "ForeheadOcculsion";
                worksheet.Cells[1, 8].Value = "Quality";
           

                // Add data to the worksheet
                int row = 2;
                foreach (var filename in facesDetected)
                {
                    worksheet.Cells[row, 1].Value = filename.Item1;
                    worksheet.Cells[row, 2].Value = "Face Detected";
                    worksheet.Cells[row, 3].Value = filename.Item2;
                    worksheet.Cells[row, 4].Value = filename.Item3;
                    worksheet.Cells[row, 5].Value = filename.Item4;
                    worksheet.Cells[row, 6].Value = filename.Item5;
                    worksheet.Cells[row, 7].Value = filename.Item6;
                    worksheet.Cells[row, 8].Value = filename.Item7;
                    row++;
                }
                foreach (var filename in facesNotDetected)
                {
                    worksheet.Cells[row, 1].Value = filename.Item1;
                    worksheet.Cells[row, 2].Value = "No Face Detected";
                    worksheet.Cells[row, 3].Value = filename.Item2;
                    worksheet.Cells[row, 4].Value = filename.Item3;
                    worksheet.Cells[row, 5].Value = filename.Item4;
                    worksheet.Cells[row, 6].Value = filename.Item5;
                    worksheet.Cells[row, 7].Value = filename.Item6;
                    worksheet.Cells[row, 8].Value = filename.Item7;
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
                 folderPath = @"C:\PPDENHANCEDIMAGES\";
            }
            else
            {
                 folderPath = @"C:\PPDPhotos\";
            }
           

            string outputFolder = @"C:\PPDENHANCEDIMAGES\";
            string secondoutputFolder = @"C:\PPDREENHANCEDIMAGES\";

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
                        image.Mutate(x => x.GaussianBlur(0.2f));
                        // Apply a brightness filter to the image

                        image.Mutate(x => x.Brightness((float)(0.2) * 5));
                    }
                    else
                    {
                        image.Mutate(x => x.GaussianSharpen(1.3f));
                        image.Mutate(x => x.GaussianBlur(0.35f));
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

