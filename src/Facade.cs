using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsx2string
{
    public static class Facade
    {
        private static List<IExporter> exportList = new List<IExporter>();

        /// <summary>
        /// 初始化注册器
        /// </summary>
        private static void IniRegister()
        {
            ReisterExporter(new JsonExporter());
            ReisterExporter(new LuaExporter());
            ReisterExporter(new SQLExporter());
            ReisterExporter(new TextExporter());

            ReisterExporter(new CplusExporter());
            ReisterExporter(new CsharpExporter());
            ReisterExporter(new GoLangExporter());
            ReisterExporter(new JavaExporter());
        }

        /// <summary>
        /// 注册导出器
        /// </summary>
        private static void ReisterExporter(IExporter exporter)
        {
            if (exporter == null) {
                return;
            }
            if (!exportList.Contains(exporter)) {
                exportList.Add(exporter);
            }
        }

        public static List<string> ProcessCore(Options options)
        {
            if (options.ExcelPath.Length <= 0) {
                return null;
            }
            if (!Directory.Exists(options.ExcelPath)) {
                return null;
            }
            string[] files = Directory.GetFiles(options.ExcelPath, "*.xlsx", SearchOption.AllDirectories);
            if (files.Length <= 0) {
                return null;
            }
            List<string> sb = new List<string>();
            foreach (string file in files) {
                options.ExcelPath = file;
                sb.Add(file);
            }
            return sb;
        }

        /// <summary>
        /// 根据命令行参数，执行Excel数据导出工作
        /// </summary>
        /// <param name="options">命令行参数</param>
        public static void ParseExcel(Options options)
        {
            // 加载Excel文件
            using (FileStream excelFile = File.Open(options.ExcelPath, FileMode.Open, FileAccess.Read)) {
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelFile);
                excelReader.IsFirstRowAsColumnNames = true;
                DataSet datSet = excelReader.AsDataSet();

                // 数据检测
                if (datSet.Tables.Count < 1) {
                    throw new Exception("Excel文件中没有找到Sheet: " + options.ExcelPath);
                }

                // 取得数据
                DataTable sheet = datSet.Tables[0];
                if (sheet.Rows.Count <= 0) {
                    throw new Exception("Excel Sheet中没有数据: " + options.ExcelPath);
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
            }
        }
    }
}
