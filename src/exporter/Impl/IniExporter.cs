using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System;

/***
 * TextExporter.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
namespace xlsx2string
{
    public class IniExporter : ExporterBase
    {
        private static Dictionary<string, StringBuilder> sbTextMap;
        private static StringBuilder sbConst;

        private static Dictionary<string, string> languagemapping = new Dictionary<string, string>() 
        {
            { "zh_cn", "ChineseSimplified" },
            { "zh_tw", "ChineseTraditional" },
            { "zh_hk", "ChineseHongkong" },
            { "en", "English" },
            { "ja", "Japanese" },
        };

        public override void New()
        {
            base.New();
            sbTextMap = new Dictionary<string, StringBuilder>();
            sbConst = new StringBuilder();

            sbConst.AppendLine(
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

            sbConst.AppendLine();

            sbConst.AppendLine("using xsj.framework;");
            sbConst.AppendLine();

            sbConst.AppendLine("namespace xsj");
            sbConst.AppendLine("{");

            sbConst.AppendLine("\tpublic static class LocalizationConst\r\n\t{");
            sbConst.AppendLine();
        }

        public override void Process(bool isOver = true)
        {
            ProcessIniText(isOver);
            ProcessIniConst(isOver);
        }

        private void ProcessIniText(bool isOver)
        {
            for (int i = 1; i < ColCount; ++i) {
                if (!sbTextMap.ContainsKey(FieldList[i].name)) {
                    sbTextMap[FieldList[i].name] = new StringBuilder();
                    FieldDef field = FieldList[i];
                    sbTextMap[FieldList[i].name].AppendLine(string.Format(";{0}", field.comment));
                    sbTextMap[FieldList[i].name].AppendLine(string.Format("[{0}]", languagemapping[field.name.ToLower()]));
                }
            }

            int filterRow = 0;
            foreach (DataRow row in Sheet.Rows) {
                if (filterRow <= 1) {
                    filterRow++;
                    continue;
                }
                object[] columns = row.ItemArray;
                for (int i = 1; i < columns.Length; ++i) {
                    sbTextMap[FieldList[i].name].AppendFormat("{0}={1}", columns[0], columns[i]);
                    sbTextMap[FieldList[i].name].AppendLine();
                }
            }

            if (isOver) {
                int index = 1;
                foreach (KeyValuePair<string, StringBuilder> pair in sbTextMap) {
                    // 写文件
                    string dirname = null;
                    if (languagemapping.TryGetValue(pair.Key.ToLower(), out dirname)) {
                        string p = string.Format("{0}/{1}/Dictionaries/{2}.ini.txt", "localization", dirname, "Default");
                        p = Path.Combine(Path.GetDirectoryName(Option.TxtPath), p);
                        StringBuilder sb = sbTextMap[FieldList[index].name];
                        this.WriteFile(p, sb.ToString(), Coding);
                    }
                    index++;
                }
            }
        }

        private void ProcessIniConst(bool isOver)
        {
            int filterRow = 0;
            foreach (DataRow row in Sheet.Rows) {
                if (filterRow <= 1) {
                    filterRow++;
                    continue;
                }
                object[] columns = row.ItemArray;
                if(columns.Length <= 0) {
                    throw new Exception("I18N excel error.");
                }
                sbConst.AppendFormat("\t\tpublic const string L_{0} = \"{1}\";", columns[0].ToString().Trim(), columns[0].ToString().Trim());
                sbConst.AppendLine();
            }

            if (isOver) {

                sbConst.AppendLine("\t}");
                sbConst.AppendLine("}");

                // 写文件
                string path = GetFileName(Option.TxtPath);
                path = string.Format("{0}/LocalizationConst.cs", "cs");
                path = Path.Combine(Path.GetDirectoryName(Option.TxtPath), path);

                // 写文件
                WriteFile(path, sbConst.ToString(), Coding);
            }
        }
    }
}
