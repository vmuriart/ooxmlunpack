namespace OoXmlUnpack
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Reflection;
    using System.Xml.Linq;

    class Program
    {
        static Dictionary<int, XElement> SharedStrings = new Dictionary<int, XElement>();
        static Dictionary<int, int> KeptSharedStrings = new Dictionary<int, int>();
 
        static void Main(string[] args)
        {
            string sourcePath = args.Length == 0
                                    ? Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                                    : args[0];

            try
            {
                ProcessPath(sourcePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error!");
                Console.WriteLine(ex.Message);
            }
        }

        private static void ProcessPath(string sourcePath)
        {
            foreach (var file in new DirectoryInfo(sourcePath).EnumerateFiles())
            {
                if (file.Extension == ".xlsx" || file.Extension == ".xlsm")
                {
                    ProcessExcelFile(file.FullName);
                }
            }

            foreach (var folder in new DirectoryInfo(sourcePath).EnumerateDirectories())
            {
                ProcessPath(folder.FullName);
            }
        }

        private static void ProcessExcelFile(string sourceFile)
        {
            var destFile = sourceFile;
            var extractFolder = sourceFile + ".extracted";
            if (Directory.Exists(extractFolder))
            {
                if (new DirectoryInfo(extractFolder).LastWriteTime == new FileInfo(sourceFile).LastAccessTime)
                {
                    return;
                }

                Directory.Delete(extractFolder, true);
            }

            Console.WriteLine("File: " + sourceFile);

            File.Copy(sourceFile, sourceFile + ".orig", true);

            ZipFile.ExtractToDirectory(sourceFile, extractFolder);

            ProcessExtractedFolder(extractFolder);

            File.Delete(destFile);
            ZipFile.CreateFromDirectory(extractFolder, destFile, CompressionLevel.NoCompression, false);

            // Make sure the modified date of the extract folder is the same as the decompressed, repacked source file
            new DirectoryInfo(extractFolder).LastWriteTime = new FileInfo(sourceFile).LastAccessTime;

            SharedStrings.Clear();
            KeptSharedStrings.Clear();
        }

        private static void ProcessExtractedFolder(string extractFolder)
        {
            foreach (var file in new DirectoryInfo(extractFolder).EnumerateFiles())
            {
                ProcessExtractedFile(file.FullName);
            }

            foreach (var folder in new DirectoryInfo(extractFolder).EnumerateDirectories())
            {
                ProcessExtractedFolder(folder.FullName);
            }
        }

        private static void ProcessExtractedFile(string fileName)
        {
            if (Path.GetFileName(fileName) == "calcChain.xml")
            {
                File.Delete(fileName);
                return;
            }
            
            XDocument doc;
            try
            {
                doc = XDocument.Load(fileName);

                Console.WriteLine(fileName);
            }
            catch (Exception ex)
            {
                return;
            }

            var ns = doc.Root.Name.Namespace;

            if (Path.GetFileName(fileName) == "sharedStrings.xml")
            {
                int sharedStringId = 0;
                int keptSharedStringId = 0;
                foreach(var sharedString in doc.Root.Elements(ns + "si").ToList())
                {
                    if (sharedString.Element(ns + "t") == null)
                    {
                        KeptSharedStrings.Add(sharedStringId, keptSharedStringId);
                        keptSharedStringId++;
                    }
                    else
                    {
                        SharedStrings.Add(sharedStringId, sharedString);
                        sharedString.Remove();
                    }

                    sharedStringId++;
                }

                doc.Root.Attribute("count").Value = "0";
                doc.Root.Attribute("uniqueCount").Value = "0";
            }

            foreach (var row in doc.Descendants(ns + "row"))
            {
                foreach(var cell in row.Elements(ns + "c"))
                {
                    var formula = cell.Element(ns + "f");
                    var value = cell.Element(ns + "v");

                    if (value != null)
                    {
                        if (formula != null)
                        {
                            value.Remove();
                        }
                        else if (cell.Attribute("t") != null && cell.Attribute("t").Value == "s")
                        {
                            var sharedStringId = int.Parse(value.Value);
                            int keptSharedStringId;
                            if (KeptSharedStrings.TryGetValue(sharedStringId, out keptSharedStringId))
                            {
                                value.Value = keptSharedStringId.ToString();
                            }
                            else
                            {
                                var sharedString = SharedStrings[sharedStringId];
                                value.Remove();
                                cell.Add(new XElement(ns + "is", sharedString.Descendants()));
                                cell.Attribute("t").Value = "inlineStr";
                            }
                        }
                    }
                }
            }

            try
            {
                doc.Save(fileName);
            }
            catch (Exception ex)
            {
                
            }
        }
    }
}
