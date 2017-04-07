using System;
using System.Collections.Generic;
using System.Data;

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

        // 导出器
        private static Dictionary<ExportType, IExporter> exporterCached;
        // 导出对象
        private static Dictionary<ExportType, List<Options>> exporterOption;
        // 导出设置参数
        private static OptionsForm exporterOptionsForm;
        // 表单数据缓存
        private static Dictionary<string, DataTable> sheetCached;

        // 导出总数
        private static int exportTotalCount = 0;

        static DataMemory()
        {
            checkerInfoCached = new Dictionary<string, string>();
            exporterInfoCached = new Dictionary<string, string>();

            exporterCached = new Dictionary<ExportType, IExporter>();
            exporterOption = new Dictionary<ExportType, List<Options>>();
            exporterOptionsForm = new OptionsForm();

            sheetCached = new Dictionary<string, DataTable>();

            IniMemory();
        }

        /// <summary>
        /// 初始化缓存
        /// </summary>
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

        /// <summary>
        /// 设置导出器
        /// </summary>
        /// <param name="type"></param>
        /// <param name="exporter"></param>
        private static void SetExporter(ExportType type, IExporter exporter)
        {
            IExporter tmpExporter = null;
            if(exporterCached.TryGetValue(type, out tmpExporter)) {
                return;
            }
            exporterCached.Add(type, exporter);
        }

        /// <summary>
        /// 获得导出容器
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
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

        /// <summary>
        /// 设置导出类型
        /// </summary>
        /// <param name="type"></param>
        public static void SetOptionFormType(ExportType type)
        {
            exporterOptionsForm.SetExportType(type);
        }

        /// <summary>
        /// 移除导出类型
        /// </summary>
        /// <param name="type"></param>
        public static void RemOptionFromType(ExportType type)
        {
            exporterOptionsForm.RemExportType(type);
        }

        /// <summary>
        /// 获得导出设置参数
        /// </summary>
        /// <returns></returns>
        public static OptionsForm GetOptionsFrom()
        {
            return exporterOptionsForm;
        }

        /// <summary>
        /// 获得所有导出类型
        /// </summary>
        /// <returns></returns>
        public static List<ExportType> GetOptionsFromTypes()
        {
            return exporterOptionsForm.ExporterList;
        }

        /// <summary>
        /// 设置导出Options
        /// </summary>
        /// <param name="type"></param>
        /// <param name="option"></param>
        public static void SetExportOption(ExportType type, Options option)
        {
            List<Options> list = null;
            if(!exporterOption.TryGetValue(type, out list)) {
                list = new List<Options>();
                exporterOption.Add(type, list);
            }
            list.Add(option);
        }

        /// <summary>
        /// 获得导出列表根据类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static List<Options> GetExportOptions(ExportType type)
        {
            List<Options> list = null;
            if (!exporterOption.TryGetValue(type, out list)) {
                return null;
            }
            return list;
        }

        public static void SetSheet(string key, DataTable data)
        {
            DataTable table = null;
            if (!sheetCached.TryGetValue(key, out table)) {
                sheetCached.Add(key, data);
            }
        }

        public static DataTable GetSheet(string key)
        {
            DataTable table = null;
            if (!sheetCached.TryGetValue(key, out table)) {
                return null;
            }
            return table;
        }

        /// <summary>
        /// 设置导出计数器
        /// </summary>
        /// <param name="count"></param>
        public static void SetExportTotalCount(int count)
        {
            exportTotalCount = count;
        }

        public static int GetExportTotalCount()
        {
            return exportTotalCount;
        }

        /// <summary>
        /// 清除导出信息
        /// </summary>
        public static void Destroy()
        {
            checkerInfoCached.Clear();
            exporterInfoCached.Clear();

            exporterOption.Clear();

            sheetCached.Clear();
        }
    }
}
