using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;

namespace xlsx2string
{
    public class ExporterAll : IExporterAll
    {
        private static Dictionary<ExportType, string> TypeMapping;

        static ExporterAll()
        {
            TypeMapping = new Dictionary<ExportType, string>(); 
        }

        public ExporterAll()
        {
            New();
        }

        public List<Options> OptionList
        {
            get;
            set;
        }

        public Encoding Coding
        {
            get;
            set;
        }

        public ExportType ExpType
        {
            get;
            set;
        }

        public string OutPath
        {
            get;
            set;
        }

        /// <summary>
        /// 获取文件名称
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        protected virtual string GetFileName(string path)
        {
            return Path.GetFileNameWithoutExtension(path);
        }

        /// <summary>
        /// 写文件
        /// </summary>
        /// <param name="path"></param>
        /// <param name="context"></param>
        /// <param name="coding"></param>
        protected virtual void WriteFile(string path, string context, Encoding coding)
        {
            string outPath = Path.GetDirectoryName(path);
            if (!Directory.Exists(outPath)) {
                Directory.CreateDirectory(outPath);
            }
            if (File.Exists(path)) {
                File.Delete(path);
            }

            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write)) {
                using (TextWriter writer = new StreamWriter(stream, coding)) {
                    writer.Write(context);
                }
            }
        }

        public void New()
        {
            TypeMapping[ExportType.cs] = "ConfigInitTab";
        }

        public void Init()
        {
            
        }

        public void Process()
        {
            if(ExpType != ExportType.cs) {
                return;
            }

            StringBuilder sb = new StringBuilder();

            sb.AppendLine(
@"// ========================================
// Copyright (c) 2017 KingSoft, All rights reserved.
// http://www.kingsoft.com
// 
// Generated Code
// 
// Date:     2017/05/24
// Author:   xiangjinbao
// Email:    xiangjinbao@kingsoft.com
// ========================================"
);

            sb.AppendLine("using xsj.framework;");
            sb.AppendLine();

            sb.AppendLine();
            sb.AppendLine("namespace xsj");
            sb.AppendLine("{");
            sb.AppendFormat("\tpublic class {0}\n", TypeMapping[ExpType]);
            sb.AppendLine("\t{");
            sb.AppendLine("\t\tpublic static void InitConfig()");
            sb.AppendLine("\t\t{");
            foreach (Options option in OptionList) {
                string name = GetFileName(Options.ConvertToString(ExpType, option));
                sb.AppendFormat("\t\t\tFacade.InitTabConf<{0}Tab, {1}TabConfig>();\n", name, name);
            }
            sb.AppendLine("\t\t}");
            sb.AppendLine("\t}");
            sb.AppendLine("}");

            string p = string.Format("{0}/{1}.{2}", OutPath, TypeMapping[ExpType], ExpType);

            // 写文件
            WriteFile(p, sb.ToString(), Coding);
        }

        public void Clear()
        {
            
        }


    }
}
