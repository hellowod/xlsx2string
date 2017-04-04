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

        // 导出Exporter
        private static Dictionary<ExportType, IExporter> exportDict;

        private static List<ExportType> exporterTypes;

        static DataMemory()
        {
            checkerInfo = new Dictionary<string, string>();
            exporterInfo = new Dictionary<string, string>();

            exportDict = new Dictionary<ExportType, IExporter>();

            exporterTypes = new List<ExportType>();
        }

        public static void SetExporter(ExportType type, IExporter exporter)
        {
            IExporter tmpExporter = null;
            if(!exportDict.TryGetValue(type, out tmpExporter)) {
                return;
            }
            exportDict.Add(type, exporter);
        }

        public static IExporter GetExporter(ExportType type)
        {
            IExporter exporter = null;
            if (exportDict.TryGetValue(type, out exporter)) {
                return exporter;
            }
            return null;
        }

        public static void SetExporterType(ExportType type)
        {
            if (!exporterTypes.Contains(type)) {
                exporterTypes.Add(type);
            }
        }

        public static List<ExportType> GetExporterTypes()
        {
            return exporterTypes;
        }
    }
}
