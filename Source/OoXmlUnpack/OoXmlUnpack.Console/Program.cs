namespace OoXmlUnpack.Console
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.IO;
    using System.Reflection;
    using System.Text;
    using System.Windows.Forms;

    public static class Program
    {
        private static readonly string SourcePath = ConfigurationManager.AppSettings["SourcePath"];

        private static readonly bool KeepBackupCopy = ConfigFlag("KeepBackupCopy", true);
        private static readonly bool ProcessExtractedFiles = ConfigFlag("ProcessExtractedFiles", true);
        private static readonly bool StripValues = ConfigFlag("StripValues", false);
        private static readonly bool CleanDataLinks = ConfigFlag("CleanDataLinks", false);
        private static readonly bool InlineStrings = ConfigFlag("InlineStrings", false);
        private static readonly bool KeepExtractedFiles = ConfigFlag("KeepExtractedFiles", true);
        private static readonly bool RelativeCellRefs = ConfigFlag("RelativeCellRefs", false);
        private static readonly bool RemoveStyles = ConfigFlag("RemoveStyles", false);
        private static readonly bool RemoveFormulaTypes = ConfigFlag("RemoveFormulaTypes", false);
        private static readonly bool CodeStyleOutput = ConfigFlag("CodeStyleOutput", false);
        private static readonly bool Quiet = ConfigFlag("Quiet", false);

        static void Main()
        {
            var sourcePath = string.IsNullOrEmpty(SourcePath) ? Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) : SourcePath;

            if (!Quiet)
            {
                var message = new StringBuilder();
                message.AppendLine(string.Format("Unpacking all Excel files found within the following path:"));
                message.AppendLine(string.Format("\tSourcePath: '{0}'", sourcePath));
                message.AppendLine(string.Format("Options are as follows:"));
                message.AppendLine(string.Format("\tKeepBackupCopy: {0}", KeepBackupCopy));
                message.AppendLine(string.Format("\tProcessExtractedFiles: {0}", ProcessExtractedFiles));
                message.AppendLine(string.Format("\tStripValues: {0}", StripValues));
                message.AppendLine(string.Format("\tCleanDataLinks: {0}", CleanDataLinks));
                message.AppendLine(string.Format("\tInlineStrings: {0}", InlineStrings));
                message.AppendLine(string.Format("\tKeepExtractedFiles: {0}", KeepExtractedFiles));
                message.AppendLine(string.Format("\tRelativeCellRefs: {0}", RelativeCellRefs));
                message.AppendLine(string.Format("\tRemoveStyles: {0}", RemoveStyles));
                message.AppendLine(string.Format("\tRemoveFormulaTypes: {0}", RemoveFormulaTypes));
                message.AppendLine(string.Format("\tCodeStyleOutput: {0}", CodeStyleOutput));
                message.AppendLine(string.Format("\tQuiet: {0}", Quiet));
                message.AppendLine(string.Format("(options can be set in the app.config file)"));
                if (MessageBox.Show(message.ToString(), "Office Open XML Unpack Utility", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
                    == DialogResult.Cancel)
                {
                    return;
                }
            }

            try
            {
                var unpack = new Unpack(
                    KeepBackupCopy,
                    ProcessExtractedFiles,
                    StripValues,
                    CleanDataLinks,
                    InlineStrings,
                    KeepExtractedFiles,
                    false,
                    RemoveStyles,
                    false,
                    RemoveFormulaTypes,
                    CodeStyleOutput,
                    RelativeCellRefs);
                ProcessPath(unpack, sourcePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error!");
                Console.WriteLine(ex.Message);
            }
        }

        private static void ProcessPath(Unpack unpack, string sourcePath)
        {
            foreach (var file in new DirectoryInfo(sourcePath).EnumerateFiles())
            {
                if (file.Extension == ".xlsx" || file.Extension == ".xlsm")
                {
                    unpack.ProcessExcelFile(file.FullName);
                }
            }

            foreach (var folder in new DirectoryInfo(sourcePath).EnumerateDirectories())
            {
                ProcessPath(unpack, folder.FullName);
            }
        }

        public static List<T> TopoSort<T>(IEnumerable<T> source, Func<T, IEnumerable<T>> getDependencies)
        {
            var sorted = new List<T>();
            var visited = new Dictionary<T, bool>();

            foreach (var item in source)
            {
                Visit(item, getDependencies, sorted, visited);
            }

            return sorted;
        }

        private static void Visit<T>(T item, Func<T, IEnumerable<T>> getDependencies, List<T> sorted, Dictionary<T, bool> visited)
        {
            bool inProgress;
            var alreadyVisited = visited.TryGetValue(item, out inProgress);

            if (alreadyVisited)
            {
                if (inProgress)
                {
                    throw new ArgumentException("Cyclic dependency found.");
                }
            }
            else
            {
                visited[item] = true;

                var dependencies = getDependencies(item);
                if (dependencies != null)
                {
                    foreach (var dependency in dependencies)
                    {
                        Visit(dependency, getDependencies, sorted, visited);
                    }
                }

                visited[item] = false;
                sorted.Add(item);
            }
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
