using System;
using System.Collections.Generic;

/***
 * DataMemory.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
namespace xlsx2string
{
    public static class DataMemory
    {
        private static Dictionary<String, String> checkerInfoCached;
        private static Dictionary<String, String> exporterInfoCached;

        private static Dictionary<ExportType, IExporter> exporterCached;
        private static Dictionary<ExportType, List<Options>> exporterOption;
        private static OptionsForm exporterOptionsForm;

        static DataMemory()
        {
            checkerInfoCached = new Dictionary<string, string>();
            exporterInfoCached = new Dictionary<string, string>();

            exporterCached = new Dictionary<ExportType, IExporter>();
            exporterOption = new Dictionary<ExportType, List<Options>>();
            exporterOptionsForm = new OptionsForm();

            IniMemory();
        }

        private static void IniMemory()
        {
            IniExporter();
        }

        /// <summary>
        /// 初始化注册器
        /// </summary>
        private static void IniExporter()
        {
            SetExporter(ExportType.json, new JsonExporter());
            SetExporter(ExportType.lua, new LuaExporter());
            SetExporter(ExportType.sql, new SQLExporter());
            SetExporter(ExportType.txt, new TextExporter());

            SetExporter(ExportType.cpp, new CplusExporter());
            SetExporter(ExportType.cs, new CsharpExporter());
            SetExporter(ExportType.go, new GoLangExporter());
            SetExporter(ExportType.java, new JavaExporter());
        }

        private static void SetExporter(ExportType type, IExporter exporter)
        {
            IExporter tmpExporter = null;
            if(exporterCached.TryGetValue(type, out tmpExporter)) {
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

        /// <summary>
        /// 设置资源路径
        /// </summary>
        /// <param name="path"></param>
        public static void SetOptionFormSrcPath(string path)
        {
            if (string.IsNullOrEmpty(path)) {
                return;
            }
            exporterOptionsForm.XlsxSrcPath = path;
        }

        /// <summary>
        /// 设置导出路径
        /// </summary>
        /// <param name="path"></param>
        public static void SetOptionFormDstPath(string path)
        {
            if (string.IsNullOrEmpty(path)) {
                return;
            }
            exporterOptionsForm.XlsxDstPath = path;
        }

        public static void SetOptionFormType(ExportType type)
        {
            exporterOptionsForm.SetExportType(type);
        }

        public static void RemOptionFromType(ExportType type)
        {
            exporterOptionsForm.RemExportType(type);
        }

        public static OptionsForm GetOptionsFrom()
        {
            return exporterOptionsForm;
        }

        public static List<ExportType> GetOptionsFromTypes()
        {
            return exporterOptionsForm.ExporterList;
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
