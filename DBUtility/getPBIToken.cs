
using IO.Swagger.Api;
using IO.Swagger.Model;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace DBUtility
{
    public class getPBIToken
    {

        //string PBIApiConfig = "http://localhost/reports/api/v2.0";
        //string PBIAccountConfig = "ECISPOWERUATTES\\vfisk";
        //string PBIPasswordConfig = "fisk@EC1";
        //string PBIApiConfig = "http://192.168.1.37/ReportS/api/v2.0";
        //string PBIAccountConfig = "fiskdev\\Administrator";
        //string PBIPasswordConfig = "fiskpwd01!";

        public string PBIApiConfig = "http://10.187.55.182/reports/api/v2.0";
        public string PBIAccountConfig = "ECISPOWERUATTES\\vfisk";
        public string PBIPasswordConfig = "fisk@EC1";


        public static string connectionString = "Data Source=ECISPOWERUATTES;Initial Catalog=Eisai_DMT;User Id=sa;Password=123456;";
        //public static string connectionString = "Data Source = .\\sqlserver2019; Initial Catalog = WeicaiTest; Integrated Security = SSPI";
        //string PBIApiConfig = "http://192.168.1.37/ReportS/api/v2.0";
        //string PBIAccountConfig = "fiskdev\\Administrator";
        //string PBIPasswordConfig = "fiskpwd01!";


        public object GetPBIServerFolder()
        { 
            string PBIApi = PBIApiConfig;
            string PBIAccount = PBIAccountConfig;
            string PBIPassword = PBIPasswordConfig;

            string ServeName = "PBIL";
            string APLURL = PBIApi;
            //var Service = BE.ReportConnectionConfiguration.FirstOrDefault(x => x.ID == Id && x.ReportServiceType == ServeName);
            //NtlmAuthenticator NAuthen = new NtlmAuthenticator(Domain + CurrentUser.UserAccount, CurrentUser.UserPwd);
            NtlmAuthenticator NAuthen = new NtlmAuthenticator(PBIAccount, PBIPassword);
            CatalogItemsApi clientCI = new CatalogItemsApi(APLURL);
            clientCI.Configuration.ApiClient.RestClient.Authenticator = NAuthen;

            string[] server = APLURL.Split('/');
            var serverReport = server[0] + "/" + server[1] + "/" + server[2] + "/" + server[3] + "/";
            List<CatalogItem> cItems = clientCI.GetCatalogItems().Value;
            //文件夹
            var cItems_Folders = cItems.Where(ci => ci.Type == CatalogItemType.Folder).ToList();
            //所有报表
            List<CatalogItem> cItems_Report = null;
            var apiurl = "";
            if (ServeName.Equals("PBIL"))
            {
                apiurl = serverReport + "powerbi";
                cItems_Report = cItems.Where(ci => ci.Type == CatalogItemType.PowerBIReport).ToList();

            }
            else if (ServeName.Equals("SSRS"))
            {
                apiurl = serverReport + "report";
                cItems_Report = cItems.Where(ci => ci.Type == CatalogItemType.Report).ToList();
            }
            List<ServerFolder> modelList = new List<ServerFolder>();

            var cItems_FoldersNull = cItems_Folders.Where(x => x.ParentFolderId == null);
            foreach (var item in cItems_FoldersNull)
            {
                ServerFolder model = new ServerFolder();
                model.Id = item.Id;
                model.ParentFolderId = item.ParentFolderId;
                model.Path = item.Path;
                model.Name = item.Name;
                model.Type = "Folder";
                model.disabled = true;
                model.child = GetChild(item.Id, cItems_Folders, cItems_Report, ServeName, apiurl);

                modelList.Add(model);
            }

            var cItems_ReportNull = cItems_Report.Where(x => x.ParentFolderId == null);
            foreach (var items in cItems_ReportNull)
            {
                ServerFolder models = new ServerFolder();
                models.Id = items.Id;
                models.ParentFolderId = items.ParentFolderId;
                models.Path = apiurl + items.Path;
                models.Name = items.Name;
                models.Type = "Report";
                models.disabled = false;
                models.ReportType = items.Type.ToString();
                models.ReportServiceType = ServeName;
                modelList.Add(models);
            }

            return modelList;
        }

        private List<ServerFolder> GetChild(Guid? id, List<CatalogItem> cItems_PowerBIReport, List<CatalogItem> cItems_Report, string serveName, string aPIUrl)
        {
            List<ServerFolder> modelLidt = new List<ServerFolder>();
            var Folder = cItems_PowerBIReport.Where(e => e.ParentFolderId == id);
            foreach (var item in Folder)
            {

                ServerFolder model = new ServerFolder();
                model.Id = item.Id;
                model.ParentFolderId = item.ParentFolderId;
                model.Path = item.Path;
                model.Name = item.Name;
                model.Type = "Folder";
                model.ReportServiceType = serveName;
                model.disabled = true;
                model.child = GetChild(item.Id, cItems_PowerBIReport, cItems_Report, serveName, aPIUrl);
                modelLidt.Add(model);
            }
            var Report = cItems_Report.Where(e => e.ParentFolderId == id);
            foreach (var item in Report)
            {
                ServerFolder model = new ServerFolder();
                model.Id = item.Id;
                model.ParentFolderId = item.ParentFolderId;
                model.Path = aPIUrl + item.Path;
                model.Name = item.Name;
                model.Type = "Report";
                model.disabled = false;
                model.ReportType = item.Type.ToString();
                model.ReportServiceType = serveName;
                modelLidt.Add(model);
            }
            return modelLidt;
        }


        public object getMaxFolderId()
        {
            string id = "";
            try
            {
                //string PBIApi = ConfigHelper.GetConfigStr("PBIAPI");
                //string PBIAccount = ConfigHelper.GetConfigStr("PBIAccount");
                //string PBIPassword = ConfigHelper.GetConfigStr("PBIPassword");
                string PBIApi = PBIApiConfig;
                string PBIAccount = PBIAccountConfig;
                string PBIPassword = PBIPasswordConfig;


                string APLURL = PBIApi;
                //var Service = BE.ReportConnectionConfiguration.FirstOrDefault(x => x.ID == Id && x.ReportServiceType == ServeName);
                //NtlmAuthenticator NAuthen = new NtlmAuthenticator(Domain + CurrentUser.UserAccount, CurrentUser.UserPwd);
                NtlmAuthenticator NAuthen = new NtlmAuthenticator(PBIAccount, PBIPassword);
                CatalogItemsApi clientCI = new CatalogItemsApi(APLURL);
                clientCI.Configuration.ApiClient.RestClient.Authenticator = NAuthen;

                string[] server = APLURL.Split('/');
                var serverReport = server[0] + "/" + server[1] + "/" + server[2] + "/" + server[3] + "/";
                List<CatalogItem> cItems = clientCI.GetCatalogItems().Value;
                //文件夹
                var cItems_Folders = cItems.Where(ci => ci.Type == CatalogItemType.Folder).ToList();

                var cItems_FoldersNull = cItems_Folders.Where(x => x.ParentFolderId == null).FirstOrDefault();

                if (cItems_FoldersNull != null)
                {
                    id = cItems_FoldersNull.Id.ToString();
                }
            }
            catch (Exception ex)
            {


            }
            return id;
        }


        //public List<string> getAllRepotId() {
        //    List<string> ReportList = new List<string>();

        //    //string PBIApi = ConfigHelper.GetConfigStr("PBIAPI");
        //    //string PBIAccount = ConfigHelper.GetConfigStr("PBIAccount");
        //    //string PBIPassword = ConfigHelper.GetConfigStr("PBIPassword");

        //    string PBIApi = PBIApiConfig;
        //    string PBIAccount = PBIAccountConfig;
        //    string PBIPassword = PBIPasswordConfig;


        //    string ServeName = "PBIL";
        //    string APLURL = PBIApi;
        //    //var Service = BE.ReportConnectionConfiguration.FirstOrDefault(x => x.ID == Id && x.ReportServiceType == ServeName);
        //    //NtlmAuthenticator NAuthen = new NtlmAuthenticator(Domain + CurrentUser.UserAccount, CurrentUser.UserPwd);
        //    NtlmAuthenticator NAuthen = new NtlmAuthenticator(PBIAccount, PBIPassword);
        //    CatalogItemsApi clientCI = new CatalogItemsApi(APLURL);
        //    clientCI.Configuration.ApiClient.RestClient.Authenticator = NAuthen;

        //    string[] server = APLURL.Split('/');
        //    var serverReport = server[0] + "/" + server[1] + "/" + server[2] + "/" + server[3] + "/";
        //    List<CatalogItem> cItems = clientCI.GetCatalogItems().Value;
        //    //文件夹
        //    var cItems_Folders = cItems.Where(ci => ci.Type == CatalogItemType.Folder).ToList();
        //    //所有报表
        //    List<CatalogItem> cItems_Report = null;
        //    var apiurl = "";
        //    if (ServeName.Equals("PBIL"))
        //    {
        //        apiurl = serverReport + "powerbi";
        //        cItems_Report = cItems.Where(ci => ci.Type == CatalogItemType.PowerBIReport).ToList();

        //    }
        //    else if (ServeName.Equals("SSRS"))
        //    {
        //        apiurl = serverReport + "report";
        //        cItems_Report = cItems.Where(ci => ci.Type == CatalogItemType.Report).ToList();
        //    }
        //    List<ServerFolder> modelList = new List<ServerFolder>();

        //    var cItems_FoldersNull = cItems_Folders.Where(x => x.ParentFolderId == null);
        //    foreach (var item in cItems_FoldersNull)
        //    {

        //        GetChildReportId(ref ReportList, item.Id, cItems_Folders, cItems_Report, ServeName, apiurl);

        //    }

        //    var cItems_ReportNull = cItems_Report.Where(x => x.ParentFolderId == null);
        //    foreach (var items in cItems_ReportNull)
        //    {
        //        ReportList.Add(items.Id.ToString());
        //    }

        //    return ReportList;

        //}

        private void GetChildReportId(ref List<string> ReportList, Guid? id, List<CatalogItem> cItems_PowerBIReport, List<CatalogItem> cItems_Report, string serveName, string aPIUrl)
        {

            var Folder = cItems_PowerBIReport.Where(e => e.ParentFolderId == id);
            foreach (var item in Folder)
            {

                GetChildReportId(ref ReportList, item.Id, cItems_PowerBIReport, cItems_Report, serveName, aPIUrl);
            }
            var Report = cItems_Report.Where(e => e.ParentFolderId == id);
            foreach (var item in Report)
            {
                ReportList.Add(item.Id.ToString());
            }
        }


        public dynamic getAllRepotId()
        {
            List<string> ReportList = new List<string>();

            string PBIApi = PBIApiConfig;
            string PBIAccount = PBIAccountConfig;
            string PBIPassword = PBIPasswordConfig;


            string ServeName = "PBIL";
            string APLURL = PBIApi;
            //var Service = BE.ReportConnectionConfiguration.FirstOrDefault(x => x.ID == Id && x.ReportServiceType == ServeName);
            //NtlmAuthenticator NAuthen = new NtlmAuthenticator(Domain + CurrentUser.UserAccount, CurrentUser.UserPwd);
            NtlmAuthenticator NAuthen = new NtlmAuthenticator(PBIAccount, PBIPassword);
            CatalogItemsApi clientCI = new CatalogItemsApi(APLURL);
            clientCI.Configuration.ApiClient.RestClient.Authenticator = NAuthen;

            string[] server = APLURL.Split('/');
            var serverReport = server[0] + "/" + server[1] + "/" + server[2] + "/" + server[3] + "/";
            List<CatalogItem> cItems = clientCI.GetCatalogItems().Value;
            //文件夹
            var cItems_Folders = cItems.Where(ci => ci.Type == CatalogItemType.Folder).ToList();
            //所有报表
            List<CatalogItem> cItems_Report = null;
            var apiurl = "";
            if (ServeName.Equals("PBIL"))
            {
                apiurl = serverReport + "powerbi";
                cItems_Report = cItems.Where(ci => ci.Type == CatalogItemType.PowerBIReport).ToList();
            }

            foreach (CatalogItem item in cItems_Report)
            {
                ReportList.Add(item.Id.ToString());
            }

            //else if (ServeName.Equals("SSRS"))
            //{
            //    apiurl = serverReport + "report";
            //    cItems_Report = cItems.Where(ci => ci.Type == CatalogItemType.Report).ToList();
            //}
            //List<ServerFolder> modelList = new List<ServerFolder>();

            //var cItems_FoldersNull = cItems_Folders.Where(x => x.ParentFolderId == null);
            //foreach (var item in cItems_FoldersNull)
            //{

            //    GetChildReportId(ref ReportList, item.Id, cItems_Folders, cItems_Report, ServeName, apiurl);

            //}

            //var cItems_ReportNull = cItems_Report.Where(x => x.ParentFolderId == null);
            //foreach (var items in cItems_ReportNull)
            //{
            //    ReportList.Add(items.Id.ToString());
            //}
            FolderAndReport OmitFolder = GetOmitFolder(cItems_Folders, cItems_Report);

            var obj = new object();

            return obj = new { ReportList, OmitFolder };

        }

        /// <summary>
        /// 获取不需要同步权限的文件夹
        /// </summary>
        /// <returns></returns>
        private FolderAndReport GetOmitFolder(List<CatalogItem> AllFolderList, List<CatalogItem> AllReportList)
        {
            //string connectionString = "Data Source=ECISPOWERBI;Initial Catalog=Eisai_DMT;User Id=service;Password=fisk@EC1;";

            DbHelperSQL dbHelper = new DbHelperSQL(connectionString);
            //List<string> FolderList = new List<string>();

            FolderAndReport OmitList = new FolderAndReport();
            string FolderSql = $"SELECT FolderID FROM [dbo].[Cfg_OmitReport] WHERE Validity='1'";
            DataTable FolderDT = dbHelper.Query(FolderSql).Tables[0];

            foreach (DataRow item in FolderDT.Rows)
            {
                string FolderID = item["FolderID"].ToString();

                OmitList.OmitFolderList.Add(FolderID);
                GetOmitChild(ref OmitList, FolderID.CastToGUID(), AllFolderList, AllReportList);
            }

            return OmitList;
        }

        private void GetOmitChild(ref FolderAndReport OmitList, Guid? id, List<CatalogItem> cItems_PowerBIReport, List<CatalogItem> cItems_Report)
        {
            var Folder = cItems_PowerBIReport.Where(e => e.ParentFolderId == id);
            foreach (var item in Folder)
            {
                OmitList.OmitFolderList.Add(item.Id.ToString());
                GetOmitChild(ref OmitList, item.Id, cItems_PowerBIReport, cItems_Report);

            }

            var Report = cItems_Report.Where(e => e.ParentFolderId == id);
            foreach (var item in Report)
            {
                OmitList.OmitReportList.Add(item.Id.ToString());
                GetOmitChild(ref OmitList, item.Id, cItems_PowerBIReport, cItems_Report);
            }
        }

    }
}
