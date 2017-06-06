using CommandLine;
using System.Collections.Generic;

/***
 * Options.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
namespace xlsx2string
{
    /// <summary>
    /// 命令行参数定义
    /// </summary>
    public class Options
    {
        [Option('e', "excel", Required = true, HelpText = "输入的excel文件路径.")]
        public string ExcelPath
        {
            get;
            set;
        }

        [Option('j', "json", Required = false, HelpText = "指定输出的json文件路径.")]
        public string JsonPath
        {
            get;
            set;
        }

        [Option('t', "txt", Required = false, HelpText = "指定输出的txt文件路径.")]
        public string TxtPath
        {
            get;
            set;
        }

        [Option('c', "csv", Required = false, HelpText = "指定输出的csv文件路径.")]
        public string CsvPath
        {
            get;
            set;
        }

        [Option('h', "html", Required = false, HelpText = "指定输出的html文件路径.")]
        public string HtmlPath
        {
            get;
            set;
        }

        [Option('l', "lua", Required = false, HelpText = "指定输出的lua数据定义代码文件路径.")]
        public string LuaPath
        {
            get;
            set;
        }

        [Option('s', "sql", Required = false, HelpText = "指定输出的sql文件路径.")]
        public string SQLPath
        {
            get;
            set;
        }

        [Option('p', "csharp", Required = false, HelpText = "指定输出的c#数据定义代码文件路径.")]
        public string CSharpPath
        {
            get;
            set;
        }

        [Option('j', "java", Required = false, HelpText = "指定输出的java数据定义代码文件路径.")]
        public string JavaPath
        {
            get;
            set;
        }

        [Option('+', "cpp", Required = false, HelpText = "指定输出的cpp数据定义代码文件路径.")]
        public string CppPath
        {
            get;
            set;
        }

        [Option('g', "go", Required = false, HelpText = "指定输出的go数据定义代码文件路径.")]
        public string GoPath
        {
            get;
            set;
        }

        [Option('l', "lowcase", Required = false, DefaultValue = false, HelpText = "字段名称自动转换为小写")]
        public bool Lowcase
        {
            get;
            set;
        }

        /// <summary>
        /// 将参数转换为对象
        /// </summary>
        /// <param name="inputPath"></param>
        /// <param name="outputPath"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static Options ConvertToOption(string inputPath, string outputPath, ExportType type)
        {
            Options option = new Options();
            option.ExcelPath = inputPath;
            switch (type) {
                case ExportType.json:
                    option.JsonPath = outputPath;
                    break;
                case ExportType.txt:
                    option.TxtPath = outputPath;
                    break;
                case ExportType.html:
                    option.HtmlPath = outputPath;
                    break;
                case ExportType.csv:
                    option.CsvPath = outputPath;
                    break;
                case ExportType.lua:
                    option.LuaPath = outputPath;
                    break;
                case ExportType.cs:
                    option.CSharpPath = outputPath;
                    break;
                case ExportType.java:
                    option.JavaPath = outputPath;
                    break;
                case ExportType.cpp:
                    option.CppPath = outputPath;
                    break;
                case ExportType.go:
                    option.GoPath = outputPath;
                    break;
                case ExportType.sql:
                    option.SQLPath = outputPath;
                    break;
                default:
                    break;
            }
            return option;
        }

        /// <summary>
        /// 将对象装换为参数
        /// </summary>
        /// <param name="type"></param>
        /// <param name="option"></param>
        /// <returns></returns>
        public static string ConvertToString(ExportType type, Options option)
        {
            string tmpStr = string.Empty;
            switch (type) {
                case ExportType.json:
                    tmpStr = option.JsonPath;
                    break;
                case ExportType.txt:
                    tmpStr = option.TxtPath;
                    break;
                case ExportType.csv:
                    tmpStr = option.CsvPath;
                    break;
                case ExportType.html:
                    tmpStr = option.HtmlPath;
                    break;
                case ExportType.lua:
                    tmpStr = option.LuaPath;
                    break;
                case ExportType.cs:
                    tmpStr = option.CSharpPath;
                    break;
                case ExportType.java:
                    tmpStr = option.JavaPath;
                    break;
                case ExportType.cpp:
                    tmpStr = option.CppPath;
                    break;
                case ExportType.go:
                    tmpStr = option.GoPath;
                    break;
                case ExportType.sql:
                    tmpStr = option.SQLPath;
                    break;
                default:
                    break;
            }
            return tmpStr;
        }
    }
}
