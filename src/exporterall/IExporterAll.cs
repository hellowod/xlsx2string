using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsx2string
{
    interface IExporterAll
    {
        List<Options> OptionList
        {
            get;
            set;
        }

        Encoding Coding
        {
            get;
            set;
        }

        ExportType ExpType
        {
            get;
            set;
        }

        string OutPath
        {
            get;
            set;
        }

        /// <summary>
        /// 实例化
        /// </summary>
        void New();

        /// <summary>
        /// 初始化
        /// </summary>
        void Init();

        /// <summary>
        /// 处理
        /// </summary>
        void Process();

        /// <summary>
        /// 清理
        /// </summary>
        void Clear();
    }
}
