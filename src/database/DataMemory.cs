using System;
using System.Collections.Generic;

namespace xlsx2string
{
    public static class DataMemory
    {
        private static Dictionary<String, String> checkerInfo;
        private static Dictionary<String, String> exporterInfo;

        private static List<ExportType> exporterTypes;

        private static Dictionary<ExportType, IExporter> exporterCached;
        private static Dictionary<ExportType, List<Options>> exporterOption;

        static DataMemory()
        {
            checkerInfo = new Dictionary<string, string>();
            exporterInfo = new Dictionary<string, string>();

            exporterTypes = new List<ExportType>();

            exporterCached = new Dictionary<ExportType, IExporter>();
            exporterOption = new Dictionary<ExportType, List<Options>>();
        }

        public static void SetExporter(ExportType type, IExporter exporter)
        {
            IExporter tmpExporter = null;
            if(!exporterCached.TryGetValue(type, out tmpExporter)) {
                return;
            }
            exporterCached.Add(type, exporter);
        }

        public static IExporter GetExporter(ExportType type)
        {
            IExporter exporter = null;
            if (exporterCached.TryGetValue(type, out exporter)) {
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

        public static void RemExportType(ExportType type)
        {
            if (exporterTypes.Contains(type)) {
                exporterTypes.Remove(type);
            }
        }

        public static void SetExportOption(ExportType type, Options option)
        {
            List<Options> list = null;
            if(!exporterOption.TryGetValue(type, out list)) {
                list = new List<Options>();
                exporterOption.Add(type, list);
            }
            list.Add(option);
        }

        public static List<Options> GetExportOptions(ExportType type)
        {
            List<Options> list = null;
            if (!exporterOption.TryGetValue(type, out list)) {
                return null;
            }
            return list;
        }
    }
}
