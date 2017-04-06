using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Threading;

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
        private static int ExportCount = 0;
        private static Action<int> m_onProgressEvent;

        /// <summary>
        /// 设置进度事件
        /// </summary>
        /// <param name="evt"></param>
        public static void SetOnProgressEvent(Action<int> evt)
        {
            m_onProgressEvent = evt;
        }

        /// <summary>
        /// 检查用户输入信息
        /// </summary>
        public static string ParseCheckerUserInput()
        {
            string error = null;
            OptionsForm optionForm = DataMemory.GetOptionsFrom();
            if (optionForm == null) {
                error = "工具底层异常,请程序检查！";
            } else if (string.IsNullOrEmpty(optionForm.XlsxSrcPath)) {
                error = "Xlsx表格路径不能为空！";
            }
            return error;

        }

        /// <summary>
        /// 检查用户输入信息
        /// </summary>
        public static string ParseExportUserInput()
        {
            string error = ParseCheckerUserInput();
            if(!string.IsNullOrEmpty(error)) {
                return error;
            }
            OptionsForm optionForm = DataMemory.GetOptionsFrom();
            if (string.IsNullOrEmpty(optionForm.XlsxDstPath)) {
                error = "文本导出路径不能为空！";
            } else if (optionForm.ExporterList.Count <= 0) {
                error = "请至少选择一种导出类型！";
            }
            return error;
        }

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
            Thread thread = new Thread(new ThreadStart(StartThread));
            thread.Start();
        }

        private static void StartThread()
        {
            int count = 0;
            List<ExportType> typeList = DataMemory.GetOptionsFromTypes();
            foreach (ExportType type in typeList) {
                List<Options> optionList = DataMemory.GetExportOptions(type);
                foreach (Options option in optionList) {
                    CmdXlsx(type, option);
                    count++;
                    float value = (float)count / ExportCount;
                    if(m_onProgressEvent != null) {
                        m_onProgressEvent((int)(value * 100));
                    }
                    Thread.Sleep(10);
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
            try {
                DataTable sheet = LoadSheet(option.ExcelPath);
                if (sheet != null) {
                    // 确定编码
                    Encoding coding = new UTF8Encoding(false);
                    // 执行导出器 
                    RunExporter(type, sheet, option, coding);
                }
            } catch (System.Exception ex) {
                throw new Exception("Excel export error: " + ex.Message);
            }
        }

        /// <summary>
        /// 加载表单（使用缓存机制提高效率）
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private static DataTable LoadSheet(string path)
        {
            DataTable sheet = null;

            sheet = DataMemory.GetSheet(path);
            if (sheet != null) {
                return sheet;
            } else {
                using (FileStream excelFile = File.Open(path, FileMode.Open, FileAccess.Read)) {
                    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelFile);
                    excelReader.IsFirstRowAsColumnNames = true;
                    DataSet datSet = excelReader.AsDataSet();

                    if (datSet.Tables.Count < 1) {
                        throw new Exception("Excel not found sheet: " + path);
                    }

                    sheet = datSet.Tables[0];
                    if (sheet.Rows.Count <= 0) {
                        throw new Exception("Excel sheet not data: " + path);
                    }
                }
                DataMemory.SetSheet(path, sheet);
                return sheet;
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
