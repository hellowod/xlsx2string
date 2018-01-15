using CommandLine;

namespace xlsx2string
{
    class OptionCmmd
    {
        [Option('i', "input", Required = true, HelpText = "输入文件路径.")]
        public string InputPath
        {
            get;
            set;
        }

        [Option('o', "output", Required = false, HelpText = "输出文件路径.")]
        public string OutputPath
        {
            get;
            set;
        }
    }
}
