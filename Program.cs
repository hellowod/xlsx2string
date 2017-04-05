using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace xlsx2string
{
    public class Program
    {
        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;

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
                WinformMode();
            }
        }

        /// <summary>
        /// 命令行模式
        /// </summary>
        /// <param name="args"></param>
        private static void ConsoleMode(string[] args)
        {
            DateTime startTime = DateTime.Now;
            Options options = new Options();
            CommandLine.Parser parser = new CommandLine.Parser((with) => {
                with.HelpWriter = Console.Error;
            });

            if (parser.ParseArgumentsStrict(args, options, () => Environment.Exit(-1))) {
                try {
                    // Facade.RunXlsx(options);
                } catch (Exception ex) {
                    Console.WriteLine("Error: " + ex.Message);
                }
            }
            TimeSpan expend = DateTime.Now - startTime;
            Console.WriteLine(
                string.Format("Expend time [{1} millisecond].", expend.Milliseconds)
            );
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
