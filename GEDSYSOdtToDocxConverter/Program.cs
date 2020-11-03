using DevExpress.XtraRichEdit;
using GEDSYSOdtToDocxConverter.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GEDSYSOdtToDocxConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            DirectoryInfo directory = new DirectoryInfo(Directory.GetCurrentDirectory());
            string target = string.Empty, convertTo = string.Empty;
            for (int i = 0; i < args.Length; i++)
            {
                //Console.WriteLine(args[i]);
                switch (args[i])
                {
                    case "--directory":
                        directory = new DirectoryInfo(@args[i + 1]);
                        break;
                    case "--target":
                        target = args[i + 1];
                        break;
                    case "--convert-to":
                        convertTo = args[i + 1];
                        break;
                    default:
                        break;
                }
            }

            //Console.WriteLine(directory.FullName + " - " + target + " - " + convertTo);

            if (!directory.Exists)
            {
                Console.WriteLine("Error: El directorio indicado no existe!");
                return;
            }

            //if (string.Empty.Equals(target) || string.Empty.Equals(convertTo)) {
            //    Console.WriteLine("Error: Debe indicar la extensión de los archivos a convertir, así como la extensión del formato al que convertirá");
            //}

            ConvertFilesODTtoDOCX(directory);
            Console.WriteLine("Conversión con Devexpress Completada.\n");
            Console.WriteLine("Operación: Conversión con LibreOffice:");
            List<string> fails = new List<string>();
            ConvertFilesToLibreofficeDOCX(directory, "docx", ref fails);
            Console.WriteLine("\n________________________________________________\n");
            if (fails.Count > 0)
            {
                Console.WriteLine("Formatos no convertidos:" + fails.Count);
            }
            else
            {
                Console.WriteLine("Todos los documentos fueron convertidos correctamente!");
            }
        }

        private static void ConvertFilesODTtoDOCX(DirectoryInfo directory)
        {
            RichEditDocumentServer server = new RichEditDocumentServer();
            String newName = string.Empty;
            server.Options.Export.Html.EmbedImages = true;

            foreach (FileInfo file in directory.GetFiles())
            {
                if (file.Extension.Equals(".odt") || file.Extension.Equals("odt"))
                {
                    server.LoadDocument(@file.FullName, DocumentFormat.OpenDocument);
                    //Console.WriteLine(file.FullName);
                    newName = file.Name.Replace(file.Extension, "");
                    //Console.WriteLine(newName);
                    server.SaveDocument(directory.FullName + Path.DirectorySeparatorChar + newName, DocumentFormat.OpenXml);
                    try
                    {
                        file.Delete();
                    }
                    catch (Exception)
                    {
                        //Do nothing
                    }
                }
            }

            foreach (DirectoryInfo dir in directory.GetDirectories())
            {
                if (!dir.Name.Equals("TEMPORALES") && !dir.Name.Equals("ERROR_LOGS"))
                {
                    ConvertFilesODTtoDOCX(dir);
                }
            }
        }

        private static void ConvertFilesToLibreofficeDOCX(DirectoryInfo directory, string documentFormat, ref List<string> failList)
        {
            string newName = string.Empty;
            int exitCodeLibreoffice;

            foreach (FileInfo file in directory.GetFiles())
            {
                if (file.Extension.Equals(""))
                {
                    //Console.WriteLine(file.FullName);                
                    newName = file.Name + "." + documentFormat;
                    Console.WriteLine("\n________________________________________________\n");
                    exitCodeLibreoffice = CommandLine.RunLibreOfficeConverter(file.FullName, directory.FullName, documentFormat);
                    if (exitCodeLibreoffice == 0)
                    {
                        Console.WriteLine(string.Format("{0} ha sido convertido a formato {1}", file.FullName, documentFormat));
                        try
                        {
                            file.Delete();
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("No pudo ser eliminado: " + file.FullName);
                        }
                    }
                    else
                    {
                        Console.WriteLine(string.Format("Error: {0} No pudo ser convertido!", file.FullName));
                        failList.Add(file.FullName);
                    }
                }

            }

            foreach (DirectoryInfo dir in directory.GetDirectories())
            {
                if (!dir.Name.Equals("TEMPORALES") && !dir.Name.Equals("ERROR_LOGS"))
                {
                    ConvertFilesToLibreofficeDOCX(dir, documentFormat, ref failList);
                }

            }
        }
    }
}
