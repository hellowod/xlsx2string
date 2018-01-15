using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

/***
 * Program.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
namespace xlsx2string
{
    /// <summary>
    /// 启动程序--程序结构太差，暂时只是实现功能
    /// </summary>
    public class Program
    {
        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;

        private const string ARGS =
            "使用： xlsx2string.exe [-options]\n" +
            "参数:\n" +
            "\t-i 输入路径 路径为数据表的根目录路径\n" +
            "\t-o 输出路径 输出路径为数据表导出路径\n" +
            "退出请按任意键退出程序";

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length > 0) {
                ConsoleMode(args);
            } else {
                //WinformMode();
                Console.WriteLine(ARGS);
                Console.Read();
            }
        }

        /// <summary>
        /// 命令行模式
        /// </summary>
        /// <param name="args"></param>
        private static void ConsoleMode(string[] args)
        {
            DateTime startTime = DateTime.Now;
            OptionCmmd options = new OptionCmmd();
            CommandLine.Parser parser = new CommandLine.Parser((with) => {
                with.HelpWriter = Console.Error;
            });

            if (parser.ParseArgumentsStrict(args, options, () => Environment.Exit(-1))) {
                try {
                    OptionsForm optionForm = new OptionsForm();
                    optionForm.XlsxSrcPath = options.InputPath;
                    optionForm.XlsxDstPath = options.OutputPath;

                    optionForm.SetExportType(ExportType.cs);
                    optionForm.SetExportType(ExportType.txt);
                    DataMemory.SetOptionForm(optionForm);

                    ExprotCallbackArgv argv = new ExprotCallbackArgv();
                    argv.OnProgressChanged = (int progress)=> {
                        //Console.WriteLine(string.Format("Export progress {0}", progress));
                    };
                    argv.OnRunChanged = (string message) => {
                        Console.WriteLine(string.Format("Export message {0}", message));
                    };

                    Facade.BeforeExporterOptionForm();
                    Facade.RunXlsxForm(argv);
                    Facade.AfterExporterOptionForm();

                } catch (Exception ex) {
                    Console.WriteLine("Error: " + ex.Message);
                }
            }
            TimeSpan expend = DateTime.Now - startTime;
            Console.WriteLine(string.Format("Expend time [{0} millisecond].", expend.Milliseconds));
        }

        /// <summary>
        /// 窗口模式
        /// </summary>
        private static void WinformMode()
        {
            IntPtr console = GetConsoleWindow();
            ShowWindow(console, SW_HIDE);
            Application.Run(new XlsxForm());
        }
    }
}
