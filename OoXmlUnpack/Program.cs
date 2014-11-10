namespace OoXmlUnpack
{
    using System;
    using System.Collections.Generic;
    using System.Dynamic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Xml.Linq;

    using Ionic.Zip;

    using CompressionLevel = Ionic.Zlib.CompressionLevel;

    public static class Program
    {
        private const bool KeepBackupCopy = false;
        private const bool ProcessExtractedFiles = false;
        private const bool StripValues = false;
        private const bool InlineStrings = false;
        private const bool KeepExtractedFiles = false;

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
            var extractDir = new DirectoryInfo(extractFolder);
            if (Directory.Exists(extractFolder))
            {
                if (extractDir.LastWriteTime == new FileInfo(sourceFile).LastAccessTime)
                {
                    return;
                }

                Directory.Delete(extractFolder, true);
            }

            Console.WriteLine("File: " + sourceFile);

            if (KeepBackupCopy)
            {
                File.Copy(sourceFile, sourceFile + ".orig", true);
            }

            using (var zipFile = new Ionic.Zip.ZipFile(sourceFile))
            {
                var compressionMethods = zipFile.Select(entry => entry.CompressionMethod).Distinct().ToList();
                if (compressionMethods.Count == 1 && compressionMethods[0] == CompressionMethod.None)
                {
                    return;
                }

                zipFile.ExtractAll(extractFolder);
            }

            if (ProcessExtractedFiles)
            {
                ProcessExtractedFolder(extractFolder);
            }

            File.Delete(destFile);

            using (var zipFile = new Ionic.Zip.ZipFile(sourceFile))
            {
                zipFile.CompressionMethod = CompressionMethod.None;
                zipFile.CompressionLevel = CompressionLevel.Level0;
                zipFile.AddSelectedFiles("*.*", extractFolder, string.Empty, true);
                zipFile.Save();
            }

            if (!KeepExtractedFiles)
            {
                Directory.Delete(extractFolder, true);
            }
            else
            {
                // Make sure the modified date of the extract folder is the same as the decompressed, repacked source file
                extractDir.LastWriteTime = new FileInfo(sourceFile).LastAccessTime;
            }

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

            if (InlineStrings)
            {
                if (Path.GetFileName(fileName) == "sharedStrings.xml")
                {
                    int sharedStringId = 0;
                    int keptSharedStringId = 0;
                    foreach (var sharedString in doc.Root.Elements(ns + "si").ToList())
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
                            if (StripValues)
                            {
                                value.Remove();
                            }
                        }
                        else if (cell.Attribute("t") != null && cell.Attribute("t").Value == "s")
                        {
                            if (InlineStrings)
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
            }

            try
            {
                doc.Save(fileName);
            }
            catch (Exception ex)
            {
                
            }
        }

        public static long DirSize(DirectoryInfo dir)
        {
            return dir.EnumerateDirectories().Select(DirSize).Sum()
                   + dir.EnumerateFiles().Select(f => f.Length).Sum();
        }
    }
}
