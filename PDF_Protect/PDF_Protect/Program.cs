using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF_Protect
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter File Path (C:\\Temp\\)");
            string filePath = Console.ReadLine();

            Console.WriteLine("Enter File Name (myDoc.pdf)");
            string fileName = Console.ReadLine();

            Console.WriteLine("Enter Search Key (ID )");
            string searchKey = Console.ReadLine();

            // ProtectFile("PDFSec.pdf");

            // getFileName("PDFSec.pdf");

            // SplitDoc("C:\\Users\\kobus.jonker\\Documents\\PDF\\", "PDFFile.pdf");

            // SplitProtect("C:\\Temp\\PDF\\", fileName, "ID ");
            SplitProtect(filePath, fileName, searchKey);
        }

        static private void SplitProtect(string path, string filename, string searchString)
        {

            var fileNames = new List<string>();

            var source = Path.Combine(path, filename);
            var dest = Path.Combine(Directory.GetCurrentDirectory(), filename);
            File.Copy(source, dest, true);

            // Open the file
            PdfDocument inputDocument = PdfReader.Open(filename, PdfDocumentOpenMode.Import);

            string name = Path.GetFileNameWithoutExtension(filename);
            for (int idx = 0; idx < inputDocument.PageCount; idx++)
            {
                // Create new document
                PdfDocument outputDocument = new PdfDocument();
                outputDocument.Version = inputDocument.Version;
                outputDocument.Info.Title = String.Format("{0}_{1}.pdf", idx + 1, inputDocument.Info.Title);
                outputDocument.Info.Creator = inputDocument.Info.Creator;

                // Add the page and save it
                var page = inputDocument.Pages[idx];

                outputDocument.AddPage(page);

                var saveFileName = String.Format("{0}_{1}.pdf", name, idx + 1);
                fileNames.Add(saveFileName);
                outputDocument.Save(saveFileName);
            }



            // Set Password
            foreach (var fName in fileNames)
            {
                var id = getFileName(fName);

                if(id == "")
                {
                    continue;
                }

                ProtectFile(fName, "Admin@123", id);
            }
           
        }

        static private string getFileName(string filenameDest)
        {
            string searchString = "ID: ";
            PdfDocument document = PdfReader.Open(filenameDest, "some text");

            var page = document.Pages[0];
            for (int index = 0; index < page.Contents.Elements.Count; index++)
            {
                PdfDictionary.PdfStream stream = page.Contents.Elements.GetDictionary(index).Stream;
                var  outputText = new PDFTextExtractor().ExtractTextFromPDFBytes(stream.Value);

                var searchStringIndex = outputText.IndexOf(searchString);
                if (searchStringIndex > -1)
                {
                    var id = outputText.Substring(searchStringIndex + 4, 13);
                    return id;
                }
            }
            return "";
        }
        static private void ProtectFile(string filenameDest, string ownerPw, string userPw)
        {
           
            // Open an existing document. Providing an unrequired password is ignored.
            PdfDocument document = PdfReader.Open(filenameDest, "some text");

            PdfSecuritySettings securitySettings = document.SecuritySettings;

            // Setting one of the passwords automatically sets the security level to 
            // PdfDocumentSecurityLevel.Encrypted128Bit.
            securitySettings.UserPassword = userPw;
            securitySettings.OwnerPassword = ownerPw;

            // Don't use 40 bit encryption unless needed for compatibility
            //securitySettings.DocumentSecurityLevel = PdfDocumentSecurityLevel.Encrypted40Bit;

            // Restrict some rights.
            securitySettings.PermitAccessibilityExtractContent = false;
            securitySettings.PermitAnnotations = false;
            securitySettings.PermitAssembleDocument = false;
            securitySettings.PermitExtractContent = false;
            securitySettings.PermitFormsFill = true;
            securitySettings.PermitFullQualityPrint = false;
            securitySettings.PermitModifyDocument = true;
            securitySettings.PermitPrint = false;

            // Save the document...
            document.Save(filenameDest);
            // ...and start a viewer.
            //Process.Start(filenameDest);
        }

        static private void SplitDoc(string path, string filename)
        {
            // Get a fresh copy of the sample PDF file
            //const string filename = "PDFFile.pdf";

            //var source = Path.Combine("C:\\Users\\kobus.jonker\\Documents\\PDF\\", filename);
            var source = Path.Combine(path, filename);
            var dest = Path.Combine(Directory.GetCurrentDirectory(), filename);
            File.Copy(source, dest, true);

            // Open the file
            PdfDocument inputDocument = PdfReader.Open(filename, PdfDocumentOpenMode.Import);

            string name = Path.GetFileNameWithoutExtension(filename);
            for (int idx = 0; idx < inputDocument.PageCount; idx++)
            {
                // Create new document
                PdfDocument outputDocument = new PdfDocument();
                outputDocument.Version = inputDocument.Version;
                outputDocument.Info.Title = String.Format("Page {0} of {1}", idx + 1, inputDocument.Info.Title);
                outputDocument.Info.Creator = inputDocument.Info.Creator;

                // Add the page and save it
                outputDocument.AddPage(inputDocument.Pages[idx]);
                outputDocument.Save(String.Format("{0} - Page {1}_tempfile.pdf", name, idx + 1));
            }
        }
    }
}
