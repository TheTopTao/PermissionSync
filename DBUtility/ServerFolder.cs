using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBUtility
{
    public class ServerFolder
    {

        public Guid? Id { get; set; }
        public Guid? ParentFolderId { get; set; }
        public string Name { get; set; }
        public Guid? ReportServerGroupId { get; set; }
        public string Path { get; set; }
        public string Type { get; set; }
        public string ReportType { get; set; }
        public string ReportServiceType { get; set; }
        public bool disabled { get; set; }
        public List<ServerFolder> child { get; set; }
    }

    public class ServerFolders
    {
        public string Id { get; set; }
        public string ParentFolderId { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public string ReportType { get; set; }
        public List<ServerFolderRpeots> child { get; set; }
    }

    public class ServerFolderRpeot
    {

        public Guid? Id { get; set; }
        public Guid? ParentFolderId { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
        public string Type { get; set; }
    }


    public class ServerFolderRpeots
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string ReportServerGroupId { get; set; }
        public string Path { get; set; }
        public string EmbedUrl { get; set; }
        public string Type { get; set; }
        public string ReportType { get; set; }
        public string ReportServiceType { get; set; }


    }

}
