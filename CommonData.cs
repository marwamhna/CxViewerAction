using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CxViewerAction
{
    static class CommonData
    {
        private static bool workingOffline = false;
        public static bool IsWorkingOffline
        {
            get { return workingOffline; }
            set { workingOffline = value; }
        }

        public static long SelectedScanId
        {
            get;
            set;
        }

        public static long ProjectId
        {
            get;
            set;
        }

        public static bool IsProjectBound
        {
            get;
            set;
        }

        public static bool IsProjectPublic
        {
            get;
            set;
        }

        public static string ProjectName
        {
            get;
            set;
        }

        public static string ProjectRootPath
        {
            get;
            set;
        }
    }
}
