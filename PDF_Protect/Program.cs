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

            if (string.IsNullOrWhiteSpace(filePath))
            {
                Console.WriteLine("File Path");
                Console.WriteLine("Press enter to exit");

                Console.ReadLine();
                return;
            }

            if(filePath.Last() != '\\')
            {
                filePath += "\\";
            }


            Console.WriteLine("Enter Destination Path (C:\\Temp\\Dest\\)");
            string destPath = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(destPath))
            {
                Console.WriteLine("Destination Path");
                Console.WriteLine("Press enter to exit");

                Console.ReadLine();
                return;
            }

            if (destPath.Last() != '\\')
            {
                destPath += "\\";
            }


            Console.WriteLine("Enter File Name (myDoc.pdf)");
            string fileName = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(fileName))
            {
                Console.WriteLine("File Name");
                Console.WriteLine("Press enter to exit");
                Console.ReadLine();

                return;
            }

            var checkFileName = fileName.ToLower();
            if (!checkFileName.Contains(".pdf"))
            {
                fileName += ".pdf";
            }


            Console.WriteLine("Enter Search Key (ID#:  )");
            string searchIdKey = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(searchIdKey))
            {
                Console.WriteLine("Search ID Key Invalid");
                Console.WriteLine("Press enter to exit");
                Console.ReadLine();
                return;
            }


            Console.WriteLine("Enter Search Name (Email:  )");
            string searchNameKey = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(searchNameKey))
            {
                Console.WriteLine("Search Name Key Invalid");
                Console.WriteLine("Press enter to exit");
                Console.ReadLine();
                return;
            }



            SplitProtect(filePath, destPath, fileName, searchIdKey, searchNameKey);

            Console.WriteLine("");
            Console.WriteLine("Files created successfully");
            Console.WriteLine("Press enter to exit");

            Console.ReadLine();
        }

        static private void SplitProtect(string path, string destPath, string filename, string searchIdKey, string searchNameKey)
        {
            
            var identifiers = new List<Identifier>();

            var source = Path.Combine(path, filename);

            if (!File.Exists(source))
            {
                Console.WriteLine("File does not exist");
                Console.WriteLine("Press enter to exit");
                Console.ReadLine();
                return;
            }
            var dest = Path.Combine(destPath, filename);
            File.Copy(source, dest, true);

            // Open the file
            // PdfDocument inputDocument = PdfReader.Open(filename, PdfDocumentOpenMode.Import);
            PdfDocument inputDocument = PdfReader.Open(dest, PdfDocumentOpenMode.Import);

            //string name = Path.GetFileNameWithoutExtension(filename);
            string name = Path.GetFileNameWithoutExtension(dest);
            for (int idx = 0; idx < inputDocument.PageCount; idx++)
            {
                // Create new document
                PdfDocument outputDocument = new PdfDocument();
                outputDocument.Version = inputDocument.Version;
                outputDocument.Info.Title = String.Format("{0}_{1}.pdf", idx + 1, inputDocument.Info.Title);
                outputDocument.Info.Creator = inputDocument.Info.Creator;

                // Add the page and save it
                var page = inputDocument.Pages[idx];

                var identifier = getIdentifier(page, searchIdKey, searchNameKey);
                if(identifier == null)
                {
                    Console.WriteLine("Search Values Invalid");
                    Console.WriteLine("Press enter to exit");
                    Console.ReadLine();
                    return;
                }
                identifiers.Add(identifier);

                outputDocument.AddPage(page);

                // var saveFileName = String.Format("{0}_{1}.pdf", name, idx + 1);
                var saveFileName = String.Format("{0}.pdf", identifier.PersonName);
                
                outputDocument.Save(destPath + "\\" + saveFileName);
            }


            using(StreamWriter w = File.AppendText(destPath + "\\log.txt"))
            {
                w.WriteLine("FileName,Password");

                // Set Password
                foreach (var identifier in identifiers)
                {
                    w.WriteLine(identifier.PersonName + ".pdf,'" + identifier.IdNumber + "'");
                    ProtectFile(destPath + "\\" + identifier.PersonName + ".pdf", "Admin@123", identifier.IdNumber);
                }

            }
           

        }


        //static private string getPersonID(PdfPage page, string searchString)
        //{
        //    for (int index = 0; index < page.Contents.Elements.Count; index++)
        //    {
        //        PdfDictionary.PdfStream stream = page.Contents.Elements.GetDictionary(index).Stream;
        //        var outputText = new PDFTextExtractor().ExtractTextFromPDFBytes(stream.Value);
        //        var searchStringIndex = outputText.IndexOf(searchString);
        //        if (searchStringIndex > -1)
        //        {
        //            var id = outputText.Substring(searchStringIndex + searchString.Length, 13);
        //            return id;
        //        }
        //    }
        //    return "";
        //}

        //static private string getPersonName(PdfPage page, string searchString)
        //{
        //    for (int index = 0; index < page.Contents.Elements.Count; index++)
        //    {
        //        PdfDictionary.PdfStream stream = page.Contents.Elements.GetDictionary(index).Stream;
        //        var outputText = new PDFTextExtractor().ExtractTextFromPDFBytes(stream.Value);
        //        var searchStringIndex = outputText.IndexOf(searchString);
        //        if (searchStringIndex > -1)
        //        {
        //            var id = outputText.Substring(searchStringIndex + searchString.Length, 13);
        //            return id;
        //        }
        //    }
        //    return "";
        //}


        static private Identifier getIdentifier(PdfPage page, string idKey, string nameKey)
        {
            Identifier identifier = new Identifier();
           
            for (int index = 0; index < page.Contents.Elements.Count; index++)
            {
                PdfDictionary.PdfStream stream = page.Contents.Elements.GetDictionary(index).Stream;
                var outputText = new PDFTextExtractor().ExtractTextFromPDFBytes(stream.Value);

                var searchIdIndex = outputText.IndexOf(idKey);
                if (searchIdIndex > -1)
                {
                    var id = outputText.Substring(searchIdIndex + idKey.Length, 13);
                    identifier.IdNumber = id;
                    if (identifier.isLoaded)
                    {
                        return identifier;
                    }
                }

                var searchNameIndex = outputText.IndexOf(nameKey);
                if (searchNameIndex > -1)
                {
                    var attIndex = outputText.IndexOf("@");

                    var startIndex = searchNameIndex + nameKey.Count();
                    var endIndex = attIndex - startIndex;
                    var name = outputText.Substring(startIndex, endIndex);
                    identifier.PersonName = name;
                    if (identifier.isLoaded)
                    {
                        return identifier;
                    }

                }

            }
            return null;
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
            securitySettings.PermitModifyDocument = false;
            securitySettings.PermitPrint = true;
            

            // Save the document...
            document.Save(filenameDest);
            // ...and start a viewer.
            //Process.Start(filenameDest);
        }

        //static private void SplitDoc(string path, string filename)
        //{
        //    // Get a fresh copy of the sample PDF file
        //    //const string filename = "PDFFile.pdf";
        //    //var source = Path.Combine("C:\\Users\\kobus.jonker\\Documents\\PDF\\", filename);
        //    var source = Path.Combine(path, filename);
        //    var dest = Path.Combine(Directory.GetCurrentDirectory(), filename);
        //    File.Copy(source, dest, true);
        //    // Open the file
        //    PdfDocument inputDocument = PdfReader.Open(filename, PdfDocumentOpenMode.Import);
        //    string name = Path.GetFileNameWithoutExtension(filename);
        //    for (int idx = 0; idx < inputDocument.PageCount; idx++)
        //    {
        //        // Create new document
        //        PdfDocument outputDocument = new PdfDocument();
        //        outputDocument.Version = inputDocument.Version;
        //        outputDocument.Info.Title = String.Format("Page {0} of {1}", idx + 1, inputDocument.Info.Title);
        //        outputDocument.Info.Creator = inputDocument.Info.Creator;
        //        // Add the page and save it
        //        outputDocument.AddPage(inputDocument.Pages[idx]);
        //        outputDocument.Save(String.Format("{0} - Page {1}_tempfile.pdf", name, idx + 1));
        //    }
        //}
    }

    public class Identifier
    {
        public string IdNumber { get; set; }
        public string PersonName { get; set; }

        public Boolean isLoaded
        {
            get
            {
                if (string.IsNullOrEmpty(IdNumber) || string.IsNullOrEmpty(PersonName))
                {
                    return false;
                }

                return true;
            }
        }

    }
}
