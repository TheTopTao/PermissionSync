using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DBUtility;
using System.Web;
using Microsoft.AnalysisServices.Core;

namespace PermissionSync
{
    class Program
    {
        public static log4net.ILog logInfo = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType.Name);
        static void Main(string[] args)
        {
            //PBIPermissions();
            //
            try
            {
                Permissions();

            }
            catch (Exception ex)
            {
                logInfo.Error(ex.Message);
                Console.WriteLine(ex.Message);
            }


        }

        //public static string connectionString = "Data Source=ECISPOWERUATTES;Initial Catalog=Eisai_DMT;User Id=sa;Password=123456;";
        //public static string connectionString = "Data Source=ECISPOWERBI;Initial Catalog=Eisai_DMT;User Id=service;Password=fisk@EC1;";
        public static string connectionString = "Data Source = .\\sqlserver2019; Initial Catalog = WeicaiTest; Integrated Security = SSPI";

        public static void Permissions()
        {
            try
            {

                #region 数据权限操作
                //string DomainAccount = "ROOT_EISAI"; //域账号
                //string Instance = "ECISPOWERBI"; //182：ECISPOWERUATTES， 181：ECISPOWERBIDEV
                //string DataBases = "Eisai_BAIM";    // 182： Eisai_BAIM ， 181 ：Eisai_BAIM
                //string tables = "V_Dim_Territory";
                //string tablesCVH = "V_Dim_TerritoryCVH";

                string DomainAccount = ""; //域账号
                string Instance = ".\\sqlserver2019";
                string DataBases = "Eisai_Sales";
                string tables = "V_Dim_Territory";
                string tablesCVH = "V_Dim_TerritoryCVH";

                string Field = "";
                //string connectionString = "Data Source = .\\sqlserver2019; Initial Catalog = WeicaiTest; Integrated Security = SSPI";
                //string connectionString = "Data Source=ECISPOWERBI;Initial Catalog=Eisai_DMT;User Id=service;Password=fisk@EC1;";
                //string connectionString = "Data Source=ECISPOWERUATES;Initial Catalog=Eisai_DMT;User Id=sa;Password=123456;";

                DbHelperSQL dbHelper = new DbHelperSQL(connectionString);
                //执行同步人员与角色同步关系存储过程
                SqlParameter[] pataMeter = new SqlParameter[] {
            new SqlParameter("@Instance",SqlDbType.NVarChar,100),
            new SqlParameter("@DataBases",SqlDbType.NVarChar,100),
            new SqlParameter("@Tables",SqlDbType.NVarChar,100),
            new SqlParameter("@TablesCVH",SqlDbType.NVarChar,100)
          };
                pataMeter[0].Value = Instance;
                pataMeter[1].Value = DataBases;
                pataMeter[2].Value = tables;
                pataMeter[3].Value = tablesCVH;
                dbHelper.RunProcedureAllTable("[usp_Sys_PermissionsSync]", pataMeter);

                //同步角色权限
                //string roleDataSql = $@"SELECT TR.RoleID,R.RoleName FROM [dbo].[Cfg_Tabular_Role_Mapping] TR LEFT JOIN [dbo].[Cfg_RoleInfo] R ON TR.RoleID=R.RoleID WHERE TR.DataType='Row'";
                string roleDataSql = $@"SELECT RoleID,RoleName,Flag FROM [dbo].[Cfg_RoleInfo] WHERE Validity='1'";
                DataTable RoleData = dbHelper.Query(roleDataSql).Tables[0];

                TOMHelper tmo = new TOMHelper(Instance, DataBases, tables);
                TOMHelper tmoCVH = new TOMHelper(Instance, DataBases, tablesCVH);

                //查找flag为1和3的角色，防止表模型里的角色删除
                string notDelRoleSql = $@"SELECT * FROM [dbo].[Cfg_RoleInfo] WHERE Flag ='1' OR Flag='3'";
                var notDelRole = dbHelper.Query(notDelRoleSql).Tables[0];

                //删除数据库中的角色 （flag 1和3不删除）
                tmo.delete_DataBase(Instance, DataBases, notDelRole);
                var modelRolesAll = tmo.GetModelRoles();

                foreach (DataRow role in RoleData.Rows)
                {
                    string roleName = role["RoleName"].ToString();
                    string roleID = role["RoleID"].ToString();
                    string roleFlag = role["Flag"].ToString();

                    if (roleFlag == "2")
                    {

                        bool TabHasRole = tmo.RoleIsHave(roleName);
                        if (TabHasRole)
                        {
                            //删除该角色下的用户，重新添加用户
                            tmo.DelUser(Instance, DataBases, roleName);
                        }


                        Console.WriteLine("flag2角色: " + roleName);
                        //通过角色id查找该角色下的用户的账号
                        string userSql = $@"SELECT U.UserAccount FROM [dbo].[Cfg_Role_User_Mapping] RU LEFT JOIN [dbo].[Cfg_UserInfo] U ON RU.UserID=U.UserInfoID WHERE RU.RoleID='{roleID}' AND U.Validity='1' AND RU.Validity = '1'";
                        Console.WriteLine(userSql);
                        DataTable userListDT = dbHelper.Query(userSql).Tables[0];
                        List<string> UserAccountList = new List<string>();
                        if (userListDT.Rows.Count > 0)
                        {
                            foreach (DataRow user in userListDT.Rows)
                            {
                                Console.WriteLine(user["UserAccount"].ToString());
                                UserAccountList.Add(user["UserAccount"].ToString());
                            }
                        }
                        //        //查找角色下的配置的字段
                        string fieldSql = $@"SELECT DISTINCT TR.Filed FROM [dbo].[Cfg_Tabular_Role_Mapping] TR LEFT JOIN [dbo].[Cfg_RoleInfo] R ON TR.RoleID=R.RoleID  WHERE TR.RoleID='{roleID}' and R.Validity = '1' and TR.Validity = '1' ";
                        DataTable FieldDT = dbHelper.Query(fieldSql).Tables[0];
                        foreach (DataRow field in FieldDT.Rows)
                        {
                            string fieldName = field["Filed"].ToString();
                            //通过字段和角色id  查找角色中的配置的数据
                            string dataSql = $@"SELECT DISTINCT TR.Data FROM [dbo].[Cfg_Tabular_Role_Mapping] TR LEFT JOIN [dbo].[Cfg_RoleInfo] R ON TR.RoleID=R.RoleID WHERE TR.DataType='Row' AND TR.Tables='{tables}' AND TR.Filed='{fieldName}' AND TR.RoleID='{roleID}' AND TR.Validity ='1'";
                            List<string> dataList = new List<string>();
                            DataTable dataDT = dbHelper.Query(dataSql).Tables[0];
                            foreach (DataRow data in dataDT.Rows)
                            {
                                dataList.Add(data["Data"].ToString());
                            }

                            if (!TabHasRole)
                            {

                                tmo.TabularSetRolePermissionRoleDoesNotExist(dataList, roleName, UserAccountList, DomainAccount, fieldName);
                            }

                            //同步TerritoryCVH数据权限
                            string dataCVHSql = $@"SELECT DISTINCT TR.Data FROM [dbo].[Cfg_Tabular_Role_Mapping] TR LEFT JOIN [dbo].[Cfg_RoleInfo] R ON TR.RoleID=R.RoleID WHERE TR.DataType='Row' AND TR.Tables='{tablesCVH}' AND TR.Filed='{fieldName}' AND TR.RoleID='{roleID}' AND TR.Validity = '1'";
                            List<string> dataCVHList = new List<string>();
                            DataTable dataCVHDT = dbHelper.Query(dataSql).Tables[0];
                            foreach (DataRow data in dataCVHDT.Rows)
                            {
                                dataCVHList.Add(data["Data"].ToString());
                            }
                            bool TabCVHHasRole = tmoCVH.RoleIsHave(roleName);
                            if (!TabCVHHasRole)
                            {
                                //tmoCVH.TabularSetRolePermissionRoleDoesNotExist(dataCVHList, roleName, UserAccountList, DomainAccount, fieldName);
                            }
                            else
                            {
                                //获取之前的表所有字段的权限除当前字段
                                var TabluarAssinglist = GetOtherTabularRoleFiledAuthorit(Instance, DataBases, tablesCVH, fieldName, roleID);
                                tmoCVH.TabularSetRolePermissionRoleExist(dataCVHList, roleName, TabluarAssinglist, fieldName);
                            }
                        }

                    }
                    else //若角色flag为1和3  则更新其角色下的用户
                    {
                        continue;
                        if (!modelRolesAll.Where(e => e.Name == roleName).Any())
                        {
                            continue;
                        }

                        Console.WriteLine("1和3角色需更新用户： " + roleName);
                        //删除该角色下的用户，重新添加用户
                        tmo.DelUser(Instance, DataBases, roleName);
                        //通过角色id查找该角色下的用户的账号
                        string userSql = $@"SELECT U.UserAccount FROM [dbo].[Cfg_Role_User_Mapping] RU LEFT JOIN [dbo].[Cfg_UserInfo] U ON RU.UserID=U.UserInfoID WHERE RU.RoleID='{roleID}' AND U.Validity='1' AND RU.Validity = '1' ";
                        Console.WriteLine(userSql);
                        DataTable userListDT = dbHelper.Query(userSql).Tables[0];
                        List<string> UserAccountList = new List<string>();
                        if (userListDT.Rows.Count > 0)
                        {
                            foreach (DataRow user in userListDT.Rows)
                            {
                                Console.WriteLine(user["UserAccount"].ToString());
                                UserAccountList.Add(user["UserAccount"].ToString());
                            }
                        }
                        if (UserAccountList.Count > 0)
                        {
                            Console.WriteLine("1和3开始重新添加里面的用户");
                            tmo.AddMember(Instance, DataBases, roleName, DomainAccount, UserAccountList);
                        }

                    }

                }
                #endregion
            }
            catch (Exception ex)
            {
                logInfo.Error("Permissions报错：" + ex.Message);
            }
        }

        public static void PBIPermissions()
        {
            #region 报表同步
            Console.WriteLine("====================");
            Console.WriteLine("开始同步报表权限");
            var obj = new SettingReportAuthority().SettingReportAuthoritys(); //同步报表权限

            Console.WriteLine(obj);
            //Console.WriteLine(obj.msg);
            Console.WriteLine("同步报表权限结束");
            #endregion

        }

        public static List<Tabular_Role_mapping> GetOtherTabularRoleFiledAuthorit(string Instance, string DataBase, string Table, string Filed, string RoleId)
        {

            //string connectionString = "Data Source = .\\sqlserver2019; Initial Catalog = WeicaiTest; Integrated Security = SSPI";
            //string connectionString = "Data Source=ECISPOWERBI;Initial Catalog=Eisai_DMT;User Id=service;Password=fisk@EC1;";

            DbHelperSQL dbHelper = new DbHelperSQL(connectionString);

            List<Tabular_Role_mapping> list = new List<Tabular_Role_mapping>();
            try
            {
                string sql = $@"SELECT * FROM Cfg_Tabular_Role_Mapping t WHERE t.RoleID = '{RoleId}' AND t.Instance = '{Instance}' AND t.DataBases = '{DataBase}' AND t.Tables = '{Table}' AND t.Filed != '{Filed}' AND t.Validity = '1' ";
                list = dbHelper.Query(sql).Tables[0].ToDataList<Tabular_Role_mapping>();
            }
            catch (Exception)
            {

                throw;
            }
            return list;

        }

    }
}
