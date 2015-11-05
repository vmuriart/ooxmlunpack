namespace OoXmlUnpack
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Windows.Forms;
    using System.Xml.Linq;

    using Ionic.Zip;

    using CompressionLevel = Ionic.Zlib.CompressionLevel;

    public static class Program
    {
        private static readonly string SourcePath = ConfigurationManager.AppSettings["SourcePath"];

        private static readonly bool KeepBackupCopy = ConfigFlag("KeepBackupCopy", true);
        private static readonly bool ProcessExtractedFiles = ConfigFlag("ProcessExtractedFiles", true);
        private static readonly bool StripValues = ConfigFlag("StripValues", false);
        private static readonly bool InlineStrings = ConfigFlag("InlineStrings", false);
        private static readonly bool KeepExtractedFiles = ConfigFlag("KeepExtractedFiles", true);
        private static readonly bool RelativeCellRefs = ConfigFlag("RelativeCellRefs", false);
        private static readonly bool RemoveStyles = ConfigFlag("RemoveStyles", false);
        private static readonly bool RemoveFormulaTypes = ConfigFlag("RemoveFormulaTypes", false);
        private static readonly bool Quiet = ConfigFlag("Quiet", false);

        static Dictionary<int, XElement> SharedStrings = new Dictionary<int, XElement>();
        static Dictionary<int, int> KeptSharedStrings = new Dictionary<int, int>();
 
        static void Main()
        {
            var sourcePath = string.IsNullOrEmpty(SourcePath)
                                 ? Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                                 : SourcePath;

            if (!Quiet)
            {
                var message = new StringBuilder();
                message.AppendLine(string.Format("Unpacking all Excel files found within the following path:"));
                message.AppendLine(string.Format("\tSourcePath: '{0}'", sourcePath));
                message.AppendLine(string.Format("Options are as follows:"));
                message.AppendLine(string.Format("\tKeepBackupCopy: {0}", KeepBackupCopy));
                message.AppendLine(string.Format("\tProcessExtractedFiles: {0}", ProcessExtractedFiles));
                message.AppendLine(string.Format("\tStripValues: {0}", StripValues));
                message.AppendLine(string.Format("\tInlineStrings: {0}", InlineStrings));
                message.AppendLine(string.Format("\tKeepExtractedFiles: {0}", KeepExtractedFiles));
                message.AppendLine(string.Format("\tRelativeCellRefs: {0}", RelativeCellRefs));
                message.AppendLine(string.Format("\tRemoveStyles: {0}", RemoveStyles));
                message.AppendLine(string.Format("\tRemoveFormulaTypes: {0}", RemoveFormulaTypes));
                message.AppendLine(string.Format("\tQuiet: {0}", Quiet));
                message.AppendLine(string.Format("(options can be set in the app.config file)"));
                if (MessageBox.Show(
                    message.ToString(),
                    "Office Open XML Unpack Utility",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Question) == DialogResult.Cancel)
                {
                    return;
                }
            }

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
                ////if (extractDir.LastWriteTime == new FileInfo(sourceFile).LastAccessTime)
                ////{
                ////    return;
                ////}

                Directory.Delete(extractFolder, true);
            }

            Console.WriteLine("File: " + sourceFile);

            if (KeepBackupCopy)
            {
                File.Copy(sourceFile, sourceFile + ".orig", true);
            }

            using (var zipFile = new ZipFile(sourceFile))
            {
                zipFile.ExtractAll(extractFolder);
            }

            if (ProcessExtractedFiles)
            {
                ProcessExtractedFolder(extractFolder);
            }

            File.Delete(destFile);

            using (var zipFile = new ZipFile(sourceFile))
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
            catch
            {
                return;
            }

            if (RelativeCellRefs || RemoveStyles || RemoveFormulaTypes || InlineStrings)
            {
                UpdateDocument(fileName, doc);
            }

            try
            {
                doc.Save(fileName);
            }
            catch
            {
            }
        }

        private static void UpdateDocument(string fileName, XDocument doc)
        {
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
            
            int previousRow = 0;
            foreach (var row in doc.Descendants(ns + "row"))
            {
                if (RelativeCellRefs && row.Attribute("r") != null)
                {
                    var currentRow = int.Parse(row.Attribute("r").Value);
                    row.Attribute("r").Value = "+" + (currentRow - previousRow);

                    previousRow = currentRow;
                }

                int previousColumn = 0;
                foreach (var cell in row.Elements(ns + "c"))
                {
                    var formula = cell.Element(ns + "f");
                    var value = cell.Element(ns + "v");

                    if (RelativeCellRefs && cell.Attribute("r") != null)
                    {
                        var currentColumn = Regex.Match(cell.Attribute("r").Value, "[A-Z]+").Groups[0].Value.Aggregate(
                            0,
                            (r, c) => r * 26 + c - '@');
                        cell.Attribute("r").Value = "+" + (currentColumn - previousColumn);

                        previousColumn = currentColumn;
                    }

                    if (RemoveStyles && cell.Attribute("s") != null)
                    {
                        cell.Attribute("s").Remove();
                    }

                    if (RemoveFormulaTypes && cell.Attributes("t") != null && cell.Elements().Count() == 1
                        && cell.Elements().Single().Name == (ns + "f"))
                    {
                        cell.Attributes("t").Remove();
                    }

                    if (formula != null)
                    {
                        if (StripValues)
                        {
                            if (value != null)
                            {
                                value.Remove();
                            }

                            if (cell.Attribute("t") != null)
                            {
                                cell.Attribute("t").Remove();
                            }
                        }
                    }
                    else if (value != null && cell.Attribute("t") != null && cell.Attribute("t").Value == "s")
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

        public static long DirSize(DirectoryInfo dir)
        {
            return dir.EnumerateDirectories().Select(DirSize).Sum()
                   + dir.EnumerateFiles().Select(f => f.Length).Sum();
        }

        private static bool ConfigFlag(string key, bool defaultValue)
        {
            var appSetting = ConfigurationManager.AppSettings[key];
            if (string.IsNullOrEmpty(appSetting))
            {
                return defaultValue;
            }

            if (appSetting.Equals("true", StringComparison.InvariantCultureIgnoreCase))
            {
                return true;
            }
            
            if (appSetting.Equals("false", StringComparison.InvariantCultureIgnoreCase))
            {
                return false;
            }
            
            throw new Exception(
                string.Format("If supplied, the app config setting '{0}' must be either 'true' or 'false'", key));
        }
    }
}
