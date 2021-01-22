using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBUtility
{
    public class FolderAndReport
    {
        public FolderAndReport()
        {
            OmitFolderList = new List<string>();
            OmitReportList = new List<string>();
        }
        //public int MyProperty { get; set; }
        public List<string> OmitFolderList { get; set; }

        public List<string> OmitReportList { get; set; }
    }
}
