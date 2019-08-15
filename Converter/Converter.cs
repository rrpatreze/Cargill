using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Converter
{
    class Converter
    {
        static void Main(string[] args)
        {
            if(args.Length != 2)
            {
                Console.WriteLine("Invalid number of arguments");
                Console.WriteLine(@"Usage: Converter C:\FilePath\fileName.xlsm C:\filePath\fileNewName.xlsx");
                Environment.Exit(1);
            }

            string fPath = args[0];
            string fNewFile = args[1];

            if (!File.Exists(fPath))
            {
                Console.WriteLine("File: " + fPath + " not found");
                Environment.Exit(1);
            }

            try
            {
                byte[] byteArray = File.ReadAllBytes(fPath);
                using (MemoryStream stream = new MemoryStream())
                {
                    stream.Write(byteArray, 0, (int)byteArray.Length);
                    using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(stream, true))
                    {
                        spreadsheetDoc.DeletePartsRecursivelyOfType<VbaDataPart>();
                        spreadsheetDoc.DeletePartsRecursivelyOfType<VbaProjectPart>();

                        //Change from template type to workbook type
                        spreadsheetDoc.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
                    }
                    File.WriteAllBytes(fNewFile, stream.ToArray());
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
                Environment.Exit(1);
            }

            Environment.Exit(0);
        }
    }
}
