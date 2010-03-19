using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace ReportingServicesBatchUpload
{
    internal class ReportingServicesUploader
    {
        public string ReportingServicesWebServiceUrl { get; set; }

        private ReportingServices2005WebService.ReportingService2005 _WebService { get; set; }

        public ReportingServicesUploader()
        {
            _WebService = new ReportingServicesBatchUpload.ReportingServices2005WebService.ReportingService2005();
            ReportingServicesWebServiceUrl = _WebService.Url;
        }

        /// <summary>
        /// Uploads new .rdl file to SSRS
        /// </summary>
        /// <param name="reportName"></param>
        /// <param name="rdlFilePath"></param>
        /// <param name="serverFolder"></param>
        /// <returns>List of warnings (null if successful with nor warnings</returns>
        public IEnumerable<string> UploadResource(string resourceName, string filePath, string serverFolder, bool overwrite)
        {
            Byte[] definition = Initialize(filePath);

            try {
                _WebService.CreateResource(resourceName, serverFolder, overwrite, definition, GetMimeTypeFromFileName(resourceName), null);
            }
            catch (System.Web.Services.Protocols.SoapException) {
                
            }
            catch (NotSupportedException) {
                return new string[] { string.Format("File type is not supported") };
            }

            return null;
        }

        /// <summary>
        /// Uploads new .rdl file to SSRS
        /// </summary>
        /// <param name="reportName"></param>
        /// <param name="rdlFilePath"></param>
        /// <param name="serverFolder"></param>
        /// <returns>List of warnings (null if successful with nor warnings</returns>
        public IEnumerable<string> UploadDataSource(string rdsFilePath, string serverFolder, bool overwrite)
        {
            _WebService.Url = ReportingServicesWebServiceUrl;
            _WebService.Credentials = System.Net.CredentialCache.DefaultCredentials;

            string name;
            var dataSource = GetDataSourceFromRds(rdsFilePath, out name);
            try {
                _WebService.CreateDataSource(name, serverFolder, overwrite, dataSource, null);
            }
            catch (System.Web.Services.Protocols.SoapException e) {
                if (e.Message.CompareTo(string.Format("The item '{0}' cannot be found. ---> The item '{0}' cannot be found.", serverFolder)) == 0) {
                    CreateFolder(serverFolder);
                    return UploadDataSource(rdsFilePath, serverFolder, overwrite);
                }
            }
            catch (NotSupportedException) {
                return new string[] { string.Format("File type is not supported") };
            }

            return null;
        }

        private ReportingServicesBatchUpload.ReportingServices2005WebService.DataSourceDefinition GetDataSourceFromRds(string rdsFilePath, out string name)
        {
            var dataSource = new ReportingServices2005WebService.DataSourceDefinition();
            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
            doc.Load(rdsFilePath);
            int propertyCount = doc.SelectNodes("/RptDataSource/ConnectionProperties/*").Count;
            if (propertyCount != 2 && propertyCount != 3)
                Console.Error.WriteLine("Please check the rds file.  There are an unexpected number of connection properties");
            dataSource.Extension = doc.SelectSingleNode("/RptDataSource/ConnectionProperties/Extension").InnerText;
            dataSource.ConnectString = doc.SelectSingleNode("/RptDataSource/ConnectionProperties/ConnectString").InnerText;            
            name = doc.SelectSingleNode("/RptDataSource/Name").InnerText;
            return dataSource;
        }

        private string GetMimeTypeFromFileName(string resourceName)
        {
            string extension = resourceName.Substring(resourceName.Length - 3);
            switch (extension) {
                case "bmp":
                    return "image/bmp";
                case "png":
                    return "image/x-png";
                case "jpg":
                default:
                    throw new NotSupportedException(string.Format("Extension {0} not supported"));
            }
        }

        /// <summary>
        /// Uploads new .rdl file to SSRS
        /// </summary>
        /// <param name="reportName"></param>
        /// <param name="rdlFilePath"></param>
        /// <param name="serverFolder"></param>
        /// <returns>List of warnings (null if successful with nor warnings</returns>
        public IEnumerable<string> UploadReport(string reportName, string rdlFilePath, string serverFolder, bool overwrite)
        {
            Byte[] definition = Initialize(rdlFilePath);

            ReportingServicesBatchUpload.ReportingServices2005WebService.Warning[] warnings = null;
            try {
                warnings = _WebService.CreateReport(reportName, serverFolder, overwrite, definition, null);
            }
            catch (System.Web.Services.Protocols.SoapException e) {
                if (e.Message.CompareTo(string.Format("The item '{0}' cannot be found. ---> The item '{0}' cannot be found.", serverFolder)) == 0) {
                    CreateFolder(serverFolder);
                    return UploadReport(reportName, rdlFilePath, serverFolder, overwrite);
                }
                if (overwrite ||
                    e.Message.CompareTo(string.Format("The item '{0}/{1}' already exists. ---> The item '{0}/{1}' already exists.",serverFolder, reportName)) != 0)
                    return new string[] { e.Detail.InnerXml.ToString() };
            }

            IList<string> rc = ProcessWarnings(warnings);

            return rc.Count != 0 ? rc : null;
        }

        private void CreateFolder(string serverFolder)
        {
            int lastSlash = serverFolder.LastIndexOfAny(new char[] { '\\', '/' });
            string folder = serverFolder.Substring(lastSlash + 1);
            string parent = serverFolder.Remove(lastSlash).Replace('\\', '/');
            parent = string.IsNullOrEmpty(parent) ? "/" : parent;
            try {
                _WebService.CreateFolder(folder, parent, null);
            }
            catch (System.Web.Services.Protocols.SoapException e) {
                if (e.Message.CompareTo(string.Format("The item '{0}' cannot be found. ---> The item '{0}' cannot be found.", parent)) == 0) {
                    CreateFolder(parent);
                    CreateFolder(serverFolder);
                }
                else
                    throw;
            }            
        }

        private static IList<string> ProcessWarnings(ReportingServicesBatchUpload.ReportingServices2005WebService.Warning[] warnings)
        {
            IList<string> rc = new List<string>();
            if (warnings != null)
                foreach (var warning in warnings)
                    rc.Add(warning.Message);
            return rc;
        }

        private Byte[] Initialize(string rdlFilePath)
        {
            _WebService.Url = ReportingServicesWebServiceUrl;
            _WebService.Credentials = System.Net.CredentialCache.DefaultCredentials;

            Byte[] definition = null;

            FileStream stream = File.OpenRead(rdlFilePath);
            definition = new Byte[stream.Length];
            stream.Read(definition, 0, (int)stream.Length);
            stream.Close();
            return definition;
        }

    }
}
