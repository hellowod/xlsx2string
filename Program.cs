using System;
using System.IO;
using System.Data;
using System.Text;
using Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace xlsx2string
{
    /// <summary>
    /// 应用程序
    /// </summary>
    sealed partial class Program
    {
        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private static void Main(string[] args)
        {
            if (args.Length > 0) {
                DateTime startTime = DateTime.Now;
                // 命令模式
                Options options = new Options();
                CommandLine.Parser parser = new CommandLine.Parser((with) => {
                    with.HelpWriter = Console.Error;
                });

                if (parser.ParseArgumentsStrict(args, options, () => Environment.Exit(-1))) {
                    // 执行导出操作
                    try {
                        Run(options);
                    } catch (Exception exp) {
                        Console.WriteLine("Error: " + exp.Message);
                    }
                }
                // 程序计时
                DateTime endTime = DateTime.Now;
                TimeSpan dur = endTime - startTime;
                Console.WriteLine(
                    string.Format("转换完成[{1}毫秒].", dur.Milliseconds)
                );
            } else {
                IntPtr console = GetConsoleWindow();
                ShowWindow(console, SW_HIDE);
                Application.Run(new ExcelForm());
            }
        }

        /// <summary>
        /// 根据命令行参数，执行Excel数据导出工作
        /// </summary>
        /// <param name="options">命令行参数</param>
        private static void Run(Options options)
        {
            string excelPath = options.ExcelPath;
            int header = options.HeaderRows;

            // 加载Excel文件
            using (FileStream excelFile = File.Open(excelPath, FileMode.Open, FileAccess.Read)) {
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelFile);
                excelReader.IsFirstRowAsColumnNames = true;
                DataSet datSet = excelReader.AsDataSet();

                // 数据检测
                if (datSet.Tables.Count < 1) {
                    throw new Exception("Excel文件中没有找到Sheet: " + excelPath);
                }

                // 取得数据
                DataTable sheet = datSet.Tables[0];
                if (sheet.Rows.Count <= 0) {
                    throw new Exception("Excel Sheet中没有数据: " + excelPath);
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
                    JsonExporter exporter = new JsonExporter(sheet, header, options.Lowcase);
                    exporter.SaveToFile(options.JsonPath, coding);
                }

                // 导出SQL文件
                if (options.SQLPath != null && options.SQLPath.Length > 0) {
                    SQLExporter exporter = new SQLExporter(sheet, header);
                    exporter.SaveToFile(options.SQLPath, coding);
                }

                // 生成C#定义文件
                if (options.CSharpPath != null && options.CSharpPath.Length > 0) {
                    string excelName = Path.GetFileName(excelPath);

                    CsharpExporter exporter = new CsharpExporter(sheet);
                    exporter.ClassComment = string.Format("// Generate From {0}", excelName);
                    exporter.SaveToFile(options.CSharpPath, coding);
                }
            }
        }
    }
}
