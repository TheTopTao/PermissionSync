using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IO.Swagger.Api;
using IO.Swagger.Model;
using RestSharp;

namespace DBUtility
{
    public class SettingReportAuthority
    {
        public static string connectionString = "Data Source=ECISPOWERUATTES;Initial Catalog=Eisai_DMT;User Id=sa;Password=123456;";
        public static string webUser = "ROOT_EISAI\\domain users";
        //public string connectionString = "Data Source = .\\sqlserver2019; Initial Catalog = WeicaiTest; Integrated Security = SSPI";

        string PBIApi = "http://10.187.55.182/reports/api/v2.0";
        string PBIAccount = "ECISPOWERUATTES\\vfisk";
        string PBIPassword = "fisk@EC1";
        //string PBIApi = "http://192.168.1.37/ReportS/api/v2.0";
        //string PBIAccount = "fiskdev\\Administrator";
        //string PBIPassword = "fiskpwd01!";

        public dynamic SettingReportAuthoritys()
        {
            getPBIToken getPBIData = new getPBIToken();

            //string PBIApi = "http://localhost/reports/api/v2.0";
            //string PBIAccount = "ECISPOWERBI\\powerbiservices";
            //string PBIPassword = "BIs!eisai2005";

            //string PBIApi = "http://192.168.1.37/ReportS/api/v2.0";
            //string PBIAccount = "fiskdev\\Administrator";
            //string PBIPassword = "fiskpwd01!";


            //string connectionString = "Data Source = .\\sqlserver2019; Initial Catalog = WeicaiTest; Integrated Security = SSPI";
            //string connectionString = "Data Source=ECISPOWERBI;Initial Catalog=Eisai_DMT;User Id=service;Password=fisk@EC1;";
            //string connectionString = "Data Source=ECISPOWERUATES;Initial Catalog=Eisai_DMT;User Id=sa;Password=123456;";
            DbHelperSQL dbHelper = new DbHelperSQL(connectionString);


            string FolderID = getPBIData.getMaxFolderId().ToString();
            Console.WriteLine("根目录文件夹："+FolderID);
            var AllData = getPBIData.getAllRepotId();
            //List<string> ReportList = getPBIData.getAllRepotId();
            List<string> ReportList = AllData.ReportList;
            FolderAndReport OmitFolder = AllData.OmitFolder;

            object obj = new { };
            Role role_ZHCN = new Role();
            role_ZHCN.Description = "可以查看文件夹、报表和订阅报表。";
            role_ZHCN.Name = "浏览者";

            Role roleM_ZHCN = new Role();
            roleM_ZHCN.Description = "可以管理报表服务器中的内容，包括文件夹、报表和资源。";
            roleM_ZHCN.Name = "内容管理员";


            Role roleS = new Role();
            roleS.Description = "可以将报表和链接报表发布到报表服务器。";
            roleS.Name = "发布者";

            Role roleD = new Role();
            roleD.Description = "可以查看报表定义。";
            roleD.Name = "报表生成器";

            //List<Role> roleList = new List<Role>();
            //List<Role> Mrole = new List<Role>();

            List<Policy> policyList = new List<Policy>();

            List<ItemPolicy> itemPolicyList = new List<ItemPolicy>();
            List<ItemPolicy> itemPolicyListFolder = new List<ItemPolicy>();


            List<Policy> AllpolicyList = new List<Policy>(); //所有的用户都加上
            try
            {
                //var ReportServerLink = BE.ReportConnectionConfiguration.Where(r => r.Validity == "1" && (r.ReportServiceType == "PBIL" || r.ReportServiceType == "SSRS")).ToList();

                NtlmAuthenticator NAuthen = new NtlmAuthenticator(PBIAccount, PBIPassword);
                CatalogItemsApi clientCI = new CatalogItemsApi(PBIApi);
                PowerBIReportsApi clientPBI = new PowerBIReportsApi(PBIApi);
                ReportsApi clientRS = new ReportsApi(PBIApi);
                FoldersApi folders = new FoldersApi(PBIApi);

                clientCI.Configuration.ApiClient.RestClient.Authenticator = NAuthen;
                clientPBI.Configuration.ApiClient.RestClient.Authenticator = NAuthen;
                clientRS.Configuration.ApiClient.RestClient.Authenticator = NAuthen;
                folders.Configuration.ApiClient.RestClient.Authenticator = NAuthen;
                //var Reports = (from p in BE.ReportConfig where p.ReportServerConfigID == itemLink.ID && p.Validity == "1" select new { p.ReportID, p.ID, p.Type }).ToList();

                //string RepotsSql = $"SELECT DISTINCT [ReportID]  FROM [dbo].[Cfg_Report_Role_Mapping]  WHERE Validity='1'";
                //var Reports = dbHelper.Query(RepotsSql).Tables[0];
                //folders.GetFolderPolicies("c3d2caf8-eb5b-4ba8-96f5-bb34549a9157");
                // List<Object> a = ;

                //folders.GetFolderPoliciesWithHttpInfo(FolderID);

                foreach (string reportID in ReportList)
                {
                    //判断是否存在于Report Server上
                    bool isAliveFolder = false;
                    try
                    {
                        var f = clientPBI.GetPowerBIReport(reportID);
                        isAliveFolder = true;
                    }
                    catch (Exception)
                    {
                        isAliveFolder = false;
                    }
                    if (!isAliveFolder)
                    {
                        continue;
                    }
                    Console.WriteLine("正在设置报表ID："+reportID);
                    //string reportID = itemReport["ReportID"].ToString();
                    string UsersSql = $@"SELECT DISTINCT U.UserAccount , U.UserInfoID  FROM [dbo].[Cfg_UserInfo] U 
                                        LEFT JOIN [dbo].[Cfg_Role_User_Mapping]  RU ON U.UserInfoID= RU.UserID
                                        LEFT JOIN [dbo].[Cfg_ReportFolder_Role_Mapping] RP ON RU.RoleID=RP.[RoleID] 
                                        LEFT JOIN [dbo].[Cfg_RoleInfo] R on R.RoleID=RP.RoleID 
                                        WHERE U.Validity='1' AND RU.Validity='1' AND RP.Validity='1' AND  
                                         R.Validity='1' AND RP.FolderID='{reportID}'";

                    List<ReportUse> users = new List<ReportUse>();
                    if (dbHelper.Query(UsersSql).Tables[0].Rows.Count > 0)
                    {
                        users = dbHelper.Query(UsersSql).Tables[0].ToDataList<ReportUse>();
                    }

                    if (users != null && users.Count > 0)
                    {


                        string pbiid = reportID;

                        foreach (ReportUse itemUser in users)
                        {

                            List<Role> roleList = new List<Role>();
                            List<Role> roleListFolder = new List<Role>();
                            //roleListFolder.Add(role_ZHCN);
                            //roleList.Add(role_ZHCN);

                            string pbiroleNameSql = $@"SELECT PBIR.FolderRoleName   FROM  [dbo].[Cfg_ReportFolder_Role_Mapping] PBIR 
                             LEFT JOIN [Cfg_Role_User_Mapping] RU ON PBIR.RoleID = RU.RoleID AND RU.Validity = '1'
                             LEFT JOIN [dbo].[Cfg_UserInfo] U ON RU.UserID=U.UserInfoID AND U.Validity='1' WHERE U.UserInfoID='{itemUser.UserInfoID}' AND PBIR.FileType='Report' AND PBIR.Validity='1'  
                             and PBIR.FolderID = '{reportID}'";

                            DataTable pbiroleNameReport = dbHelper.Query(pbiroleNameSql).Tables[0];

                            foreach (DataRow safeName in pbiroleNameReport.Rows)
                            {
                                string name = safeName["FolderRoleName"].ToString();
                                switch (name)
                                {
                                    case "内容管理员":
                                        if (!roleList.Exists(e => e.Name == roleM_ZHCN.Name))
                                        {
                                            roleList.Add(roleM_ZHCN);
                                        }
                                        break;
                                    case "报表生成器":
                                        if (!roleList.Exists(e => e.Name == roleD.Name))
                                        {
                                            roleList.Add(roleD);
                                        }
                                        break;
                                    case "浏览者":
                                        if (!roleList.Exists(e => e.Name == role_ZHCN.Name))
                                        {
                                            roleList.Add(role_ZHCN);
                                        }
                                        break;
                                }
                            }


                            Policy pol = new Policy();
                            pol.GroupUserName = itemUser.UserAccount;
                            pol.Roles = roleList;
                            policyList.Add(pol);




                            //清空角色
                            //roleList.Clear(); 
                            //roleListFolder.Clear();
                        }
                        List<Role> Mrole = new List<Role>();
                        Mrole.Add(roleM_ZHCN);

                        Policy mpol = new Policy();
                        mpol.GroupUserName = PBIAccount;
                        mpol.Roles = Mrole;

                        policyList.Add(mpol);
                        ItemPolicy itemp = new ItemPolicy("00000000-0000-0000-0000-000000000000");
                        itemp.Policies = policyList;
                        itemp.InheritParentPolicy = false;
                        itemPolicyList.Add(itemp);

                        if (!OmitFolder.OmitReportList.Exists(e => e == reportID))
                        {
                            clientPBI.SetPowerBIReportPolicies(reportID, itemPolicyList); //设置报表权限
                        }



                        //clientRS.SetReportPolicies(reportID, itemPolicyList);

                        //清空list
                        //roleList.Clear();
                        Mrole.Clear();
                        policyList.Clear();
                        itemPolicyList.Clear();
                        //roleListFolder.Clear();

                    }
                    //else
                    //{
                    //    List<Role> Mrole = new List<Role>();
                    //    //Mrole.Add(roleM_ZHCN);

                    //    policyList.Clear();
                    //    itemPolicyList.Clear();

                    //    Mrole.Add(roleM_ZHCN);

                    //    Policy mpol = new Policy();
                    //    mpol.GroupUserName = PBIAccount;
                    //    mpol.Roles = Mrole;

                    //    policyList.Add(mpol);
                    //    ItemPolicy itemp = new ItemPolicy("00000000-0000-0000-0000-000000000000");
                    //    itemp.Policies = policyList;
                    //    itemp.InheritParentPolicy = false;
                    //    itemPolicyList.Add(itemp);

                    //    if (!OmitFolder.OmitReportList.Exists(e => e == reportID))
                    //    {
                    //        clientPBI.SetPowerBIReportPolicies(reportID, itemPolicyList); //设置报表权限
                    //    }
                    //    //clientPBI.SetPowerBIReportPolicies(reportID, itemPolicyList); //设置报表权限

                    //    Mrole.Clear();
                    //    policyList.Clear();
                    //    itemPolicyList.Clear();
                    //}
                }

                //设置子文件夹权限
                string sonFolderSql = $" SELECT DISTINCT FolderID from [dbo].[Cfg_ReportFolder_Role_Mapping] where FileType='Folder' AND FolderID !='MaxID' AND Validity='1' AND FolderID IS NOT NULL";
                DataTable sonFolderData = dbHelper.Query(sonFolderSql).Tables[0];


                policyList.Clear();
                itemPolicyList.Clear();

                foreach (DataRow item in sonFolderData.Rows)
                {
                    string PBIFolderID = item["FolderID"].ToString();


                    //判断是否存在于Report Server上
                    bool isAliveFolder = false;
                    try
                    {
                        var f = folders.GetFolder(PBIFolderID);
                        isAliveFolder = true;
                    }
                    catch (Exception)
                    {
                        isAliveFolder = false;
                    }
                    if (!isAliveFolder)
                    {
                        continue;
                    }
                    Console.WriteLine("正在设置子文件夹ID：" + PBIFolderID);
                    string folderUsersSql = $@" SELECT DISTINCT U.UserAccount , U.UserInfoID FROM  [dbo].[Cfg_ReportFolder_Role_Mapping] PBIR LEFT JOIN [Cfg_Role_User_Mapping] RU ON PBIR.RoleID = RU.RoleID AND RU.Validity = '1'
                             LEFT JOIN [dbo].[Cfg_UserInfo] U ON RU.UserID=U.UserInfoID AND U.Validity='1' WHERE  PBIR.FileType='Folder' AND PBIR.Validity='1' and PBIR.FolderID='{PBIFolderID}' and U.UserAccount is not null";

                    List<ReportUse> folderUsers = new List<ReportUse>();
                    if (dbHelper.Query(folderUsersSql).Tables[0].Rows.Count > 0)
                    {
                        folderUsers = dbHelper.Query(folderUsersSql).Tables[0].ToDataList<ReportUse>();
                    }

                    if (folderUsers != null && folderUsers.Count > 0)
                    {

                        string pbiid = PBIFolderID;



                        foreach (ReportUse itemUser in folderUsers)
                        {
                            List<Role> roleListFolder = new List<Role>();
                            List<Role> roleListFolderMAX = new List<Role>();
                            //roleListFolder.Add(role_ZHCN);
                            //roleListFolderMAX.Add(role_ZHCN);


                            string pbiroleNameSql = $@"
                            SELECT PBIR.FolderRoleName   FROM  [dbo].[Cfg_ReportFolder_Role_Mapping] PBIR 
                             LEFT JOIN [Cfg_Role_User_Mapping] RU ON PBIR.RoleID = RU.RoleID AND RU.Validity = '1'
                             LEFT JOIN [dbo].[Cfg_UserInfo] U ON RU.UserID=U.UserInfoID AND U.Validity='1' WHERE U.UserInfoID='{itemUser.UserInfoID}' AND PBIR.FileType='Folder' AND PBIR.Validity='1' and PBIR.FolderID='{PBIFolderID}'";
                            DataTable pbiroleNameFolder = dbHelper.Query(pbiroleNameSql).Tables[0];

                            foreach (DataRow safeName in pbiroleNameFolder.Rows)
                            {
                                string name = safeName["FolderRoleName"].ToString();
                                switch (name)
                                {
                                    case "内容管理员":
                                        if (!roleListFolder.Exists(e => e.Name == roleM_ZHCN.Name))
                                        {
                                            roleListFolder.Add(roleM_ZHCN);
                                        }
                                        break;
                                    case "发布者":

                                        if (!roleListFolder.Exists(e => e.Name == roleS.Name))
                                        {
                                            roleListFolder.Add(roleS);
                                        }
                                        break;
                                    case "浏览者":
                                        if (!roleListFolder.Exists(e => e.Name == role_ZHCN.Name))
                                        {
                                            roleListFolder.Add(role_ZHCN);
                                        }
                                        break;
                                }
                            }


                            Policy pol = new Policy();
                            pol.GroupUserName = itemUser.UserAccount;
                            pol.Roles = roleListFolder;
                            policyList.Add(pol);


                            //roleListFolder.Clear();
                            //roleListFolderMAX.Clear();
                        }

                        List<Role> Mrole = new List<Role>();
                        Mrole.Add(roleM_ZHCN);

                        Policy mpol = new Policy();
                        mpol.GroupUserName = PBIAccount;
                        mpol.Roles = Mrole;

                        policyList.Add(mpol);
                        ItemPolicy itemp = new ItemPolicy("00000000-0000-0000-0000-000000000000");
                        itemp.Policies = policyList;
                        itemp.InheritParentPolicy = false;
                        itemPolicyList.Add(itemp);

                        if (!OmitFolder.OmitFolderList.Exists(e => e == PBIFolderID))
                        {
                            folders.SetFolderPolicies(PBIFolderID, itemPolicyList); //设置子文件夹权限
                        }

                        //clientRS.SetReportPolicies(reportID, itemPolicyList);

                        //清空list
                        //roleList.Clear();
                        Mrole.Clear();
                        policyList.Clear();
                        itemPolicyList.Clear();
                    }
                    //else
                    //{
                    //    List<Role> Mrole = new List<Role>();
                    //    //Mrole.Clear();
                    //    policyList.Clear();
                    //    Mrole.Add(roleM_ZHCN);

                    //    Policy mpol = new Policy();
                    //    mpol.GroupUserName = PBIAccount;
                    //    mpol.Roles = Mrole;

                    //    policyList.Add(mpol);
                    //    ItemPolicy itemp = new ItemPolicy("00000000-0000-0000-0000-000000000000");
                    //    itemp.Policies = policyList;
                    //    itemp.InheritParentPolicy = false;
                    //    itemPolicyList.Add(itemp);

                    //    if (!OmitFolder.OmitFolderList.Exists(e => e == PBIFolderID))
                    //    {
                    //        folders.SetFolderPolicies(PBIFolderID, itemPolicyList); //设置子文件夹权限
                    //    }

                    //    Mrole.Clear();
                    //    policyList.Clear();
                    //    itemPolicyList.Clear();
                    //}

                }


                //设置站点权限

                #region 设置最大权限
                string SelectUserSql = @" select DISTINCT U.UserAccount,U.UserInfoID from [dbo].[Cfg_ReportFolder_Role_Mapping] 
                                              RR LEFT JOIN[dbo].[Cfg_RoleInfo] R ON RR.RoleID = R.RoleID AND R.Validity = '1'
                                              LEFT JOIN[dbo].[Cfg_Role_User_Mapping] RU ON R.RoleID=RU.RoleID AND RU.Validity='1' 
                                              LEFT JOIN[dbo].[Cfg_UserInfo] U ON RU.UserID=U.UserInfoID
                                              WHERE RR.Validity='1' AND R.Validity='1' AND RU.Validity= '1' AND U.Validity= '1' and RR.FolderID= 'MaxId'";

                List<ReportUse> usersMaxId = new List<ReportUse>();
                if (dbHelper.Query(SelectUserSql).Tables[0].Rows.Count > 0)
                {
                    usersMaxId = dbHelper.Query(SelectUserSql).Tables[0].ToDataList<ReportUse>();
                }
                if (usersMaxId != null && usersMaxId.Count > 0)
                {

                    foreach (ReportUse itemUser in usersMaxId)
                    {
                        List<Role> roleList = new List<Role>();
                        List<Role> roleListFolder = new List<Role>();
                        //roleListFolder.Add(role_ZHCN);
                        //roleList.Add(role_ZHCN);



                        string pbiroleNameSql = $@"
                            SELECT PBIR.FolderRoleName  FROM  [dbo].[Cfg_ReportFolder_Role_Mapping] PBIR 
                             LEFT JOIN [Cfg_Role_User_Mapping] RU ON PBIR.RoleID = RU.RoleID AND RU.Validity = '1'
                             LEFT JOIN [dbo].[Cfg_UserInfo] U ON RU.UserID=U.UserInfoID AND U.Validity='1' WHERE U.UserInfoID='{itemUser.UserInfoID}' AND PBIR.FileType='Folder'　AND PBIR.Validity='1' AND 
                              FolderID = 'MaxID' ";

                        DataTable pbiroleNameFolder = dbHelper.Query(pbiroleNameSql).Tables[0];
                        if (pbiroleNameFolder.Rows.Count == 0)
                        {
                            continue;
                        }
                        foreach (DataRow safeName in pbiroleNameFolder.Rows)
                        {
                            string name = safeName["FolderRoleName"].ToString();
                            switch (name)
                            {
                                case "内容管理员":
                                    if (!roleListFolder.Exists(e => e.Name == roleM_ZHCN.Name))
                                    {
                                        roleListFolder.Add(roleM_ZHCN);
                                    }
                                    break;
                                case "发布者":
                                    if (!roleListFolder.Exists(e => e.Name == roleS.Name))
                                    {
                                        roleListFolder.Add(roleS);
                                    }
                                    break;
                                    //case "浏览者":
                                    //    if (!roleListFolder.Exists(e => e.Name == role_ZHCN.Name))
                                    //    {
                                    //        roleListFolder.Add(role_ZHCN);
                                    //    }
                                    //    break;
                            }
                        }
                        //roleListFolder.Add(roleS);//
                        
                        Policy polFolder = new Policy();
                        polFolder.GroupUserName = itemUser.UserAccount;
                        polFolder.Roles = roleListFolder;

                        if (AllpolicyList.Count == 0)
                        {
                            if (pbiroleNameFolder.Rows.Count != 0)
                            {
                                AllpolicyList.Add(polFolder);
                            }
                        }

                        if (!AllpolicyList.Exists(e => e.GroupUserName == polFolder.GroupUserName))
                        {
                            if (pbiroleNameFolder.Rows.Count != 0)
                            {
                                AllpolicyList.Add(polFolder);
                            }
                            // AllpolicyList.Add(polFolder);
                        }

                    }
                    //clientRS.SetReportPolicies(reportID, itemPolicyList);

                    //清空list
                    //roleList.Clear();
                    List<Role> Mrole = new List<Role>();
                    Mrole.Clear();
                    itemPolicyList.Clear();
                    Mrole.Add(roleM_ZHCN);
                    Policy mpol = new Policy();
                    mpol.GroupUserName = PBIAccount;
                    mpol.Roles = Mrole;
                    AllpolicyList.Add(mpol);

                    //给网站默认添加浏览者
                    List<Role> webRole = new List<Role>();
                    webRole.Clear();
                    webRole.Add(role_ZHCN);
                    Policy webMpol = new Policy();
                    webMpol.GroupUserName = webUser;
                    webMpol.Roles = webRole;
                    AllpolicyList.Add(webMpol);

                    ItemPolicy itempFolder = new ItemPolicy("00000000-0000-0000-0000-000000000000");
                    itempFolder.Policies = AllpolicyList;
                    itempFolder.InheritParentPolicy = false;
                    itemPolicyListFolder.Add(itempFolder);

                    folders.SetFolderPolicies(FolderID, itemPolicyListFolder); //设置最大文件夹权限
                }

                #endregion
                Console.WriteLine("报表全部同步完成");
                obj = new { success = true, msg = "配置成功" };
            }
            catch (Exception ex)
            {
                Console.WriteLine("报表同步失败"+ex.Message);
                //log.Error("报表权限同步失败,详细：" + ex);
                obj = new { success = false, msg = ex.Message };
            }


            return obj;
        }

    }
}
