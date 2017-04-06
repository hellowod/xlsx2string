using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

/***
 * Facade.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
namespace xlsx2string
{
    public static class Facade
    {
        public static int ExportCount = 0;

        /// <summary>
        /// 处理检查选项参数信息
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        public static void BeforeCheckerOptionForm()
        {
            OptionsForm form = DataMemory.GetOptionsFrom();

            if (form.XlsxSrcPath.Length <= 0) {
                return;
            }
            if (!Directory.Exists(form.XlsxSrcPath)) {
                return;
            }
        }

        /// <summary>
        /// 检查运行器
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        public static void RunCheckerXlsx()
        {
            
        }

        /// <summary>
        /// 处理检查选项参数信息
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        public static void AfterCheckerOptionForm()
        {
            OptionsForm form = DataMemory.GetOptionsFrom();

            if (form.XlsxSrcPath.Length <= 0) {
                return;
            }
            if (!Directory.Exists(form.XlsxSrcPath)) {
                return;
            }
        }

        /// <summary>
        /// 处理导出选项参数前事件
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        public static void BeforeExporterOptionForm()
        {
            ExportCount = 0;

            OptionsForm optionForm = DataMemory.GetOptionsFrom();
            List<ExportType> typeList = DataMemory.GetOptionsFromTypes();

            if(typeList.Count <= 0) {
                return;
            }
            if (optionForm.XlsxSrcPath.Length <= 0) {
                return;
            }
            if(optionForm.XlsxDstPath.Length <= 0) {
                return;
            }
            if (!Directory.Exists(optionForm.XlsxSrcPath)) {
                return;
            }

            string[] files = Directory.GetFiles(optionForm.XlsxSrcPath, "*.xlsx", SearchOption.AllDirectories);
            if (files.Length <= 0) {
                return;
            }
            // 注意xlsx文件命名规则： 标号_英文名_中文名
            foreach (string srcFile in files) {
                string fileName = string.Empty;
                string xlsxName = Path.GetFileNameWithoutExtension(srcFile);
                string[] xlsxNameArray = xlsxName.Split('_');
                if(xlsxNameArray.Length > 1) {
                    fileName = xlsxNameArray[1];
                } else {
                    fileName = xlsxNameArray[0];
                }
                foreach(ExportType type in typeList) {
                    string outFileName = string.Format("{0}.{1}", fileName, type);
                    string dstFile = Path.Combine(optionForm.XlsxDstPath, type.ToString(), outFileName);
                    Options option = Options.Convert(srcFile, dstFile, type);
                    DataMemory.SetExportOption(type, option);
                    ExportCount++;
                }
            }
        }

        /// <summary>
        /// 根据窗口参数，执行Excel数据导出工作
        /// </summary>
        /// <param name="options">命令行参数</param>
        public static void RunXlsxForm()
        {
            List<ExportType> typeList = DataMemory.GetOptionsFromTypes();
            foreach(ExportType type in typeList) {
                List<Options> optionList = DataMemory.GetExportOptionsByType(type);
                foreach(Options option in optionList) {
                    CmdXlsx(type, option);
                }
            }
        }

        /// <summary>
        /// 处理导出选项参数后事件
        /// </summary>
        public static void AfterExporterOptionForm()
        {

        }

        /// <summary>
        /// 根据命令行参数，执行Excel数据导出工作(函数性能需要优化)
        /// </summary>
        /// <param name="type"></param>
        /// <param name="option"></param>
        public static void CmdXlsx(ExportType type, Options option)
        {
            // 加载Excel文件
            using (FileStream excelFile = File.Open(option.ExcelPath, FileMode.Open, FileAccess.Read)) {
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelFile);
                excelReader.IsFirstRowAsColumnNames = true;
                DataSet datSet = excelReader.AsDataSet();

                // 数据检测
                if (datSet.Tables.Count < 1) {
                    throw new Exception("Excel not found sheet: " + option.ExcelPath);
                }

                // 取得数据
                DataTable sheet = datSet.Tables[0];
                if (sheet.Rows.Count <= 0) {
                    throw new Exception("Excel sheet not data: " + option.ExcelPath);
                }

                // 确定编码
                Encoding coding = new UTF8Encoding(false);
                if (option.Encoding != "utf8-nobom") {
                    foreach (EncodingInfo info in Encoding.GetEncodings()) {
                        Encoding e = info.GetEncoding();
                        if (e.EncodingName == option.Encoding) {
                            coding = e;
                            break;
                        }
                    }
                }

                 // 执行导出器 
                RunExporter(type, sheet, option, coding);
            }
        }

        /// <summary>
        /// 执行导出器
        /// </summary>
        /// <param name="type"></param>
        /// <param name="sheet"></param>
        /// <param name="option"></param>
        /// <param name="coding"></param>
        private static void RunExporter(ExportType type, DataTable sheet, Options option, Encoding coding)
        {
            IExporter exporter = DataMemory.GetExporter(type);
            if (exporter == null) {
                return;
            }
            exporter.Sheet = sheet;
            exporter.Option = option;
            exporter.Coding = coding;

            exporter.Init();
            exporter.Export();
            exporter.Clear();
        }
    }
}
