﻿using Excel;
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
        // 确定编码
        public static Encoding Coding = new UTF8Encoding(false);
        private static int languageCount = 0;

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
            
        }

        /// <summary>
        /// 检查运行器
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        public static void RunCheckerXlsx(CheckeCallbackArgv argv)
        {
            if (argv == null) {
                throw new Exception("Run xlsx form argv is null.");
            }

            if (argv.OnRunChanged != null) {
                argv.OnRunChanged("=================检查开始=================");
            }

            OptionsForm optionForm = DataMemory.GetOptionsFrom();
            if (optionForm.XlsxSrcPath.Length <= 0) {
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
            foreach (string file in files) {
                string fileName = string.Empty;
                string xlsxOutExtName = Path.GetFileNameWithoutExtension(file);
                string xlsxWithExtName = Path.GetFileName(file);
                string[] xlsxNameArray = xlsxOutExtName.Split('_');
                if (xlsxNameArray.Length <= 1) {
                    if(argv.OnRunChanged != null) {
                        argv.OnRunChanged(string.Format("{0}    name error", xlsxWithExtName));
                    }
                } else {
                    if (argv.OnRunChanged != null) {
                        argv.OnRunChanged(string.Format("{0}    ok", xlsxWithExtName));
                    }
                }
                Thread.Sleep(10);
            }
            if (argv.OnRunChanged != null) {
                argv.OnRunChanged("=================检查完毕=================");
            }
        }

        /// <summary>
        /// 处理检查选项参数信息
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        public static void AfterCheckerOptionForm()
        {
            
        }

        /// <summary>
        /// 处理导出选项参数前事件
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        public static void BeforeExporterOptionForm()
        {
            DataMemory.SetExportTotalCount(0);

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
            int count = 0;
            foreach (string srcFile in files) {
                string fileName = string.Empty;
                string xlsxName = Path.GetFileNameWithoutExtension(srcFile);
                fileName = xlsxName;
                foreach (ExportType type in typeList) {
                    string outFileName = string.Format("{0}.{1}", fileName, type);
                    string dstFile = Path.Combine(optionForm.XlsxDstPath, outFileName);
                    Options option = Options.ConvertToOption(srcFile, dstFile, type);
                    DataMemory.SetExportOption(type, option);
                    count++;
                }
            }

            DataMemory.SetExportTotalCount(count);
        }

        /// <summary>
        /// 根据窗口参数，执行Excel数据导出工作
        /// </summary>
        /// <param name="options">命令行参数</param>
        public static void RunXlsxForm(ExprotCallbackArgv argv)
        {
            if(argv == null) {
                throw new Exception("Run xlsx form argv is null.");
            }

            if (argv.OnRunChanged != null) {
                argv.OnRunChanged("=================开始导出=================");
            }

            int curreExportCount = 0;
            List<ExportType> typeList = DataMemory.GetOptionsFromTypes();
            foreach (ExportType type in typeList) {
                List<Options> optionList = DataMemory.GetExportOptions(type);
                foreach (Options option in optionList) {
                    Options result = CmdXlsx(type, option);
                    if(argv.OnRunChanged != null) {
                        argv.OnRunChanged(Options.ConvertToString(type, result));
                    }
                    curreExportCount++;
                    if (argv.OnProgressChanged != null) {
                        argv.OnProgressChanged(curreExportCount);
                    }
                    Thread.Sleep(1);
                }
            }

            RunExporterAll();

            if (argv.OnRunChanged != null) {
                argv.OnRunChanged("=================导出完毕=================");
            }
        }

        /// <summary>
        /// 处理导出选项参数后事件
        /// </summary>
        public static void AfterExporterOptionForm()
        {
            DataMemory.Destroy();
        }

        /// <summary>
        /// 根据命令行参数，执行Excel数据导出工作
        /// </summary>
        /// <param name="type"></param>
        /// <param name="option"></param>
        /// <returns></returns>
        public static Options CmdXlsx(ExportType type, Options option)
        {
            if(string.IsNullOrEmpty(Options.ConvertToString(type, option))) {
                return null;
            }
            try {
                DataTable sheet = LoadSheet(option.ExcelPath);
                if (sheet != null) {
                    RunExporter(type, sheet, option);
                }
            } catch (System.Exception ex) {
                throw new Exception("Excel table error: " + option.ExcelPath + ex.Message + ex.StackTrace);
            }

            return option;
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
                try
                {
                    FileStream stream  = File.Open(path, FileMode.Open, FileAccess.Read);
                    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    excelReader.IsFirstRowAsColumnNames = true;
                    DataSet datSet = excelReader.AsDataSet();

                    if (datSet.Tables.Count < 1) {
                        throw new Exception("Excel not found sheet: " + path);
                    }

                    sheet = datSet.Tables[0];
                    if (sheet.Rows.Count <= 0) {
                        throw new Exception("Excel sheet not data: " + path);
                    }

                    DataMemory.SetSheet(path, sheet);

                    excelReader.Close();
                } catch (System.Exception ex) {
                    throw new Exception(ex.StackTrace);
                }
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
        private static void RunExporter(ExportType type, DataTable sheet, Options option)
        {
            IExporter exporter = null;

            string path = Options.ConvertToString(type, option);
            if (path.Contains(Config.LANGUAGE)) {
                if(type == ExportType.txt) {
                    languageCount++;
                    exporter = DataMemory.GetExporter(ExportType.i18n);
                    if (exporter != null) {
                        exporter.Sheet = sheet;
                        exporter.Option = option;
                        exporter.Coding = Coding;

                        exporter.Init();
                        if(languageCount == DataMemory.GetExportLanguageTotalCount()) {
                            exporter.Process();
                        } else {
                            exporter.Process(false);
                        }
                        exporter.Clear();
                    }
                }

            } else {
                exporter = DataMemory.GetExporter(type);
                if (exporter != null) {
                    exporter.Sheet = sheet;
                    exporter.Option = option;
                    exporter.Coding = Coding;

                    exporter.Init();
                    exporter.Process();
                    exporter.Clear();
                }
            }
        }

        private static void RunExporterAll()
        {

            List<ExportType> typeList = DataMemory.GetOptionsFromTypes();
            foreach (ExportType type in typeList) {
                ExporterAll all = new ExporterAll();
                all.OptionList = DataMemory.GetExportOptions(type);
                all.Coding = Coding;
                all.ExpType = type;
                all.OutPath = GetExporterPath(type);

                all.Init();
                all.Process();
                all.Clear();
            }
        }

        private static string GetExporterPath(ExportType type)
        {
            return string.Format("{0}/{1}", DataMemory.GetOptionsFrom().XlsxDstPath, type.ToString());
        }
    }
}
