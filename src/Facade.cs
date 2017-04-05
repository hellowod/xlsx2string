using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace xlsx2string
{
    public static class Facade
    {
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
                    string dstFile = Path.Combine(optionForm.XlsxDstPath, outFileName);
                    Options option = Options.Convert(srcFile, dstFile, type);
                    DataMemory.SetExportOption(type, option);
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
                List<Options> optionList = DataMemory.GetExportOptions(type);
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
        /// 根据命令行参数，执行Excel数据导出工作
        /// </summary>
        /// <param name="type"></param>
        /// <param name="options"></param>
        public static void CmdXlsx(ExportType type, Options options)
        {
            // 加载Excel文件
            using (FileStream excelFile = File.Open(options.ExcelPath, FileMode.Open, FileAccess.Read)) {
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelFile);
                excelReader.IsFirstRowAsColumnNames = true;
                DataSet datSet = excelReader.AsDataSet();

                // 数据检测
                if (datSet.Tables.Count < 1) {
                    throw new Exception("Excel not found sheet: " + options.ExcelPath);
                }

                // 取得数据
                DataTable sheet = datSet.Tables[0];
                if (sheet.Rows.Count <= 0) {
                    throw new Exception("Excel sheet not data: " + options.ExcelPath);
                }

                // 确定编码
                Encoding coding = new UTF8Encoding(false);
                if (options.Encoding != "utf8-nobom") {
                    foreach (EncodingInfo info in Encoding.GetEncodings()) {
                        Encoding e = info.GetEncoding();
                        if (e.EncodingName == options.Encoding) {
                            coding = e;
                            break;
                        }
                    }
                }

                IExporter exporter = DataMemory.GetExporter(type);
                if(exporter != null) {
                    exporter.ToFile(sheet, options, coding);
                }

                /*
                // 导出JSON文件
                if (options.JsonPath != null && options.JsonPath.Length > 0) {
                    JsonExporter exporter = new JsonExporter(sheet, options.HeaderRows, options.Lowcase);
                    exporter.SaveToFile(options.JsonPath, coding);
                }

                // 导出SQL文件
                if (options.SQLPath != null && options.SQLPath.Length > 0) {
                    SQLExporter exporter = new SQLExporter(sheet, options.HeaderRows);
                    exporter.SaveToFile(options.SQLPath, coding);
                }

                // 生成C#定义文件
                if (options.CSharpPath != null && options.CSharpPath.Length > 0) {
                    string excelName = Path.GetFileName(options.ExcelPath);

                    CsharpExporter exporter = new CsharpExporter(sheet);
                    exporter.ClassComment = string.Format("// Generate From {0}", excelName);
                    exporter.SaveToFile(options.CSharpPath, coding);
                }
                */
            }
        }
    }
}
