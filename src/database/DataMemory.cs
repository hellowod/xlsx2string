using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsx2string
{
    public static class DataMemory
    {
        private static Dictionary<String, String> checkerInfo;
        private static Dictionary<String, String> exporterInfo;

        static DataMemory()
        {
            checkerInfo = new Dictionary<string, string>();
            exporterInfo = new Dictionary<string, string>();
        }
    }
}
