using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO.Compression;

namespace MergeDocumentsToPDF
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> Districts = new List<string>();
            args = new string[] { "all" };
            if (args.Length > 0)
            {
                string districtName = args[0];
                Districts.Add(districtName);

                if (districtName.Equals("all"))
                {
                    string DistrictsDirectory = ConfigurationManager.AppSettings["DistrictsDirectory"];
                    Districts = Directory.EnumerateDirectories(DistrictsDirectory).ToList();
                }
            }
            
            foreach (var d in Districts)
            {
                Console.WriteLine("==========================================================");
                Console.WriteLine("Started District : " + d);
                var Tehsils = Directory.EnumerateDirectories(d);
                foreach (var t in Tehsils)
                {
                    string tehsilName = t.Split("\\".ToCharArray())[t.Split("\\".ToCharArray()).Length - 1];
                    Console.WriteLine("==========================================================");
                    Console.WriteLine("Started Tehsil : " + tehsilName);
                    var Docs = Directory.GetFiles(t, "*.*", SearchOption.AllDirectories).OrderBy(f => f);
                    List<string> tehsilPdfFiles = new List<string>();
                    foreach (var doc in Docs)
                    {
                        tehsilPdfFiles.Add(GeneratePDFIntrop(doc));
                    }

                    var TehsilFileName = t + ".pdf";
                    Console.WriteLine("Merging to PDF : " + TehsilFileName);
                    CombineMultiplePDFs(tehsilPdfFiles, TehsilFileName);

                    try
                    {
                        foreach (var f in tehsilPdfFiles)
                        {
                            File.Delete(f);
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                string zipFile = string.Format("{0}.zip", d);
                GenerateZip(d, zipFile);

            }

            Console.WriteLine("Completed.");
            Console.ReadKey();
        }

        private static string GeneratePDFSpire(string doc)
        {
            string pdfFileName = doc.Split(".".ToCharArray())[0] + ".pdf";
            Spire.Doc.Document document = new Spire.Doc.Document();
            document.LoadFromFile(doc);
            document.SaveToFile(pdfFileName, Spire.Doc.FileFormat.PDF);
            Console.WriteLine("Generated : " + pdfFileName);
            return pdfFileName;
        }

        public static string GeneratePDFIntrop(string doc)
        {
            string pdfFileName = Path.GetDirectoryName(doc) + "\\" + Path.GetFileNameWithoutExtension(doc) + ".pdf";
            try
            {
                Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document wordDocument = appWord.Documents.Open(doc);
                appWord.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize;
                appWord.Visible = false;
                wordDocument.ExportAsFixedFormat(pdfFileName, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                //wordDocument.Close();
                //appWord.Documents.Close();
                appWord.Quit();
                if (wordDocument != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument);
                if (appWord != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(appWord);
                wordDocument = null;
                appWord = null;
                GC.Collect(); // force final cleanup!
                Console.WriteLine("Pdf Generated : " + pdfFileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in pdf Generation : " + pdfFileName + Environment.NewLine + ex.Message);
                // Delete pdf
                try { if (File.Exists(pdfFileName)) File.Delete(pdfFileName); }
                catch (Exception) { }
                return string.Empty;
            }
            return pdfFileName;
        }

        public static void GenerateZip(string DistrictDirectory, string zipFilePath)
        {
            var Docs = Directory.GetFiles(DistrictDirectory, "*.pdf", SearchOption.TopDirectoryOnly).OrderBy(f => f);
            using (ZipArchive zipFile = ZipFile.Open(zipFilePath, ZipArchiveMode.Create))
            {
                foreach (var doc in Docs)
                {
                    var FileName = new FileInfo(doc).Name;
                    zipFile.CreateEntryFromFile(doc, FileName, CompressionLevel.Fastest);
                }
            }
            Console.WriteLine("Zip Completed.");
        }

        public static void CombineMultiplePDFs(List<string> fileNames, string outFile)
        {
            // step 1: creation of a document-object
            Document document = new Document();

            // step 2: we create a writer that listens to the document
            PdfCopy writer = new PdfCopy(document, new FileStream(outFile, FileMode.Create));
            if (writer == null)
            {
                return;
            }

            // step 3: we open the document
            document.Open();

            foreach (string fileName in fileNames)
            {
                try
                {
                    // we create a reader for a certain document
                    PdfReader reader = new PdfReader(fileName);
                    reader.ConsolidateNamedDestinations();

                    // step 4: we add content
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        PdfImportedPage page = writer.GetImportedPage(reader, i);
                        writer.AddPage(page);
                    }

                    reader.Close();
                }
                catch (Exception)
                {
                    // Any error in reading PDF file, source files corrupted, delete it
                    if (File.Exists(fileName))
                        File.Delete(fileName);

                    continue;
                }
            }

            // step 5: we close the document and writer
            writer.Close();
            document.Close();
        }

    }
}
