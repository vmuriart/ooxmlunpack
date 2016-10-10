// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Unpack.cs" company="Ariadne Insight Ltd">
//   All code copyright 2012-2016 Ariadne Insight Ltd. All rights reserved
// </copyright>
// <summary>
//   Defines the Unpack type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace OoXmlUnpack
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Xml.Linq;

    using Ionic.Zip;
    using Ionic.Zlib;

    using OoXmlUnpack.Graph;

    public class Unpack
    {
        private readonly bool keepBackupCopy;
        private readonly bool processExtractedFiles;
        private readonly bool stripValues;
        private readonly bool cleanDataLinks;
        private readonly bool inlineStrings;
        private readonly bool keepExtractedFiles;
        private readonly bool removeCellRefs;
        private readonly bool removeStyles;
        private readonly bool removeRowNumbers;
        private readonly bool removeFormulaTypes;
        private readonly bool codeStyleOutput;
        private readonly bool relativeCellRefs;
        private readonly Dictionary<int, XElement> sharedStrings = new Dictionary<int, XElement>();
        private readonly Dictionary<int, int> keptSharedStrings = new Dictionary<int, int>();

        public Unpack(
            bool keepBackupCopy = false,
            bool processExtractedFiles = true,
            bool stripValues = false,
            bool cleanDataLinks = false,
            bool inlineStrings = false,
            bool keepExtractedFiles = true,
            bool removeCellRefs = false,
            bool removeStyles = false,
            bool removeRowNumbers = false,
            bool removeFormulaTypes = false,
            bool codeStyleOutput = false,
            bool relativeCellRefs = false)
        {
            this.keepBackupCopy = keepBackupCopy;
            this.processExtractedFiles = processExtractedFiles;
            this.stripValues = stripValues;
            this.cleanDataLinks = cleanDataLinks;
            this.inlineStrings = inlineStrings;
            this.keepExtractedFiles = keepExtractedFiles;
            this.removeCellRefs = removeCellRefs;
            this.removeStyles = removeStyles;
            this.removeRowNumbers = removeRowNumbers;
            this.removeFormulaTypes = removeFormulaTypes;
            this.codeStyleOutput = codeStyleOutput;
            this.relativeCellRefs = relativeCellRefs;
        }

        public void ProcessExcelFile(string sourceFile, string extractFolder = null)
        {
            if (extractFolder == null)
            {
                extractFolder = sourceFile + ".extracted";
            }

            var extractDir = new DirectoryInfo(extractFolder);
            if (Directory.Exists(extractFolder))
            {
                Directory.Delete(extractFolder, true);
            }

            Console.WriteLine("File: " + sourceFile);

            if (this.keepBackupCopy)
            {
                File.Copy(sourceFile, sourceFile + ".orig", true);
            }

            using (var zipFile = new ZipFile(sourceFile))
            {
                zipFile.ExtractAll(extractFolder);
            }

            if (this.processExtractedFiles)
            {
                this.ProcessExtractedFolder(new DirectoryInfo(extractFolder));
            }

            File.Delete(sourceFile);

            using (var zipFile = new ZipFile(sourceFile))
            {
                zipFile.CompressionMethod = CompressionMethod.None;
                zipFile.CompressionLevel = CompressionLevel.Level0;
                zipFile.AddSelectedFiles("*.*", extractFolder, string.Empty, true);
                zipFile.Save();
            }

            if (!this.keepExtractedFiles)
            {
                Directory.Delete(extractFolder, true);
            }
            else
            {
                // Make sure the modified date of the extract folder is the same as the decompressed, repacked source file
                var zipFile = new FileInfo(sourceFile);
                this.SetExtractedFileDates(
                    new DirectoryInfo(extractFolder),
                    zipFile.CreationTime,
                    zipFile.LastWriteTime,
                    zipFile.LastAccessTime);
            }

            this.sharedStrings.Clear();
            this.keptSharedStrings.Clear();
        }

        private void SetExtractedFileDates(DirectoryInfo directoryInfo, DateTime creationTime, DateTime lastWriteTime, DateTime lastAccessTime)
        {
            directoryInfo.CreationTime = creationTime;
            directoryInfo.LastWriteTime = lastWriteTime;
            directoryInfo.LastAccessTime = lastAccessTime;

            foreach (var file in directoryInfo.EnumerateFiles())
            {
                file.CreationTime = creationTime;
                file.LastWriteTime = lastWriteTime;
                file.LastAccessTime = lastAccessTime;
            }

            foreach (var subDirectory in directoryInfo.EnumerateDirectories())
            {
                this.SetExtractedFileDates(subDirectory, creationTime, lastWriteTime, lastAccessTime);
            }
        }

        private void ProcessExtractedFolder(DirectoryInfo extractFolder)
        {
            // This has to be UTC otherwise changes in locale or daylight savings will cause spurious changes in the zip archive
            var fileDateTime = new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Utc);

            foreach (var file in extractFolder.EnumerateFiles())
            {
                this.ProcessExtractedFile(file);

                file.Refresh();
                if (file.Exists)
                {
                    file.CreationTime = fileDateTime;
                    file.LastAccessTime = fileDateTime;
                    file.LastWriteTime = fileDateTime;
                }
            }

            foreach (var folder in extractFolder.EnumerateDirectories())
            {
                this.ProcessExtractedFolder(folder);
            }

            extractFolder.Refresh();
            if (extractFolder.Exists)
            {
                extractFolder.CreationTime = fileDateTime;
                extractFolder.LastAccessTime = fileDateTime;
                extractFolder.LastWriteTime = fileDateTime;
            }
        }

        private void ProcessExtractedFile(FileInfo file)
        {
            if (file.Name == "calcChain.xml")
            {
                file.Delete();
                return;
            }

            if (this.codeStyleOutput && Regex.IsMatch(file.Name, @"sheet\d+.xml") && file.Directory.Name == "worksheets")
            {
                new SheetToGraph(file).ConvertToCode();
                return;
            }

            XDocument doc;
            try
            {
                doc = XDocument.Load(file.FullName);

                Console.WriteLine(file);
            }
            catch
            {
                return;
            }

            if (this.removeRowNumbers || this.removeCellRefs || this.removeStyles || this.removeFormulaTypes || this.inlineStrings)
            {
                this.UpdateDocument(file, doc);
            }

            try
            {
                doc.Save(file.FullName);
            }
            catch
            {
            }
        }

        private void UpdateDocument(FileInfo file, XDocument doc)
        {
            var ns = doc.Root.Name.Namespace;
            if (this.inlineStrings)
            {
                if (file.Name == "sharedStrings.xml")
                {
                    int sharedStringId = 0;
                    int keptSharedStringId = 0;
                    foreach (var sharedString in doc.Root.Elements(ns + "si").ToList())
                    {
                        if (sharedString.Element(ns + "t") == null)
                        {
                            this.keptSharedStrings.Add(sharedStringId, keptSharedStringId);
                            keptSharedStringId++;
                        }
                        else
                        {
                            this.sharedStrings.Add(sharedStringId, sharedString);
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
                if (this.relativeCellRefs && row.Attribute("r") != null)
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

                    if (this.relativeCellRefs && cell.Attribute("r") != null)
                    {
                        var currentColumn = Regex.Match(cell.Attribute("r").Value, "[A-Z]+").Groups[0].Value.Aggregate(
                            0,
                            (r, c) => r * 26 + c - '@');
                        cell.Attribute("r").Value = "+" + (currentColumn - previousColumn);

                        previousColumn = currentColumn;
                    }

                    if (this.removeStyles && cell.Attribute("s") != null)
                    {
                        cell.Attribute("s").Remove();
                    }

                    if (this.removeFormulaTypes && cell.Attributes("t") != null && cell.Elements().Count() == 1
                        && cell.Elements().Single().Name == (ns + "f"))
                    {
                        cell.Attributes("t").Remove();
                    }

                    if (formula != null)
                    {
                        if (this.stripValues)
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
                        if (this.inlineStrings)
                        {
                            var sharedStringId = int.Parse(value.Value);
                            int keptSharedStringId;
                            if (this.keptSharedStrings.TryGetValue(sharedStringId, out keptSharedStringId))
                            {
                                value.Value = keptSharedStringId.ToString();
                            }
                            else
                            {
                                var sharedString = this.sharedStrings[sharedStringId];
                                value.Remove();
                                cell.Add(new XElement(ns + "is", sharedString.Descendants()));
                                cell.Attribute("t").Value = "inlineStr";
                            }
                        }
                    }
                }
            }
        }
    }
}