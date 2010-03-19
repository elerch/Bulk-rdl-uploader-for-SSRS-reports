using System;
using System.Collections.Generic;
using System.Text;

namespace ReportingServicesBatchUpload
{
    class Program
    {
        static void Main(string[] args)
        {
            // path to process, server folder, overwrite
            if (args.Length < 3 || args.Length > 4 || (args.Length == 4 && string.Compare(args[3], "/o", true) != 0)) {
                Console.Error.WriteLine("usage: ReportingServicesBatchUpload <path to RDL files> <reporting services server name> <server folder name> [/o]");
                Environment.Exit(1);
            }

            bool overwrite = args.Length == 4;
            string directory = args[0];
            string serverUrl = args[1];
            string serverFolder = args[2];            

            if (!System.IO.Directory.Exists(directory)) {
                Console.Error.WriteLine("Directory does not exist.  Exiting.  Directory specified: {0}", directory);
                Environment.Exit(1);
            }

            Console.WriteLine("Uploading files in directory '{0}' and publishing to reporting server {1}", directory, serverUrl);
            Console.WriteLine("New reports will be placed in '{0}' folder.  Overwrite is {1}", serverFolder, overwrite ? "ON" : "off");

            ProcessDirectory(directory, serverFolder, serverUrl, overwrite);
            Console.WriteLine("Processing complete");
            Environment.Exit(0);
        }

        private static void ProcessDirectory(string directory, string serverFolder, string url, bool overwrite)
        {
            var uploader = new ReportingServicesUploader();
            uploader.ReportingServicesWebServiceUrl = string.Format("{0}", url);

            IList<System.IO.FileInfo> files = new List<System.IO.FileInfo>();
            System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(directory);
            foreach (var file in dir.GetFiles("*.rds"))
                files.Add(file);
            foreach (var file in dir.GetFiles())
                if (file.Extension.ToLowerInvariant() != ".rds")
                    files.Add(file);

            foreach (var file in files) {
                Console.WriteLine("Processing file {0}", file.Name);
                string reportName = file.Name.Remove(file.Name.Length - 4);
                IEnumerable<string> warnings = null;
                switch (file.Extension.ToLowerInvariant()) {
                    case ".bmp":
                    case ".jpg":
                    case ".png":
                        warnings = uploader.UploadResource(file.Name, file.FullName, serverFolder, overwrite);
                        break;
                    case ".rds":
                        warnings = uploader.UploadDataSource(file.FullName, serverFolder, overwrite);
                        break;
                    case ".rdl":
                        warnings = uploader.UploadReport(file.Name.Remove(file.Name.Length - 4), file.FullName, serverFolder, overwrite);
                        break;
                    default:
                        Console.WriteLine("\tFile type unknown.  Skipping");
                        break;
                }

                if (warnings != null)
                    foreach (var warning in warnings)
                        Console.Error.WriteLine("\tWarning: {0}", warning);
            }
        }

    }
}
