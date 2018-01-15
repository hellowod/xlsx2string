using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using System.Text.RegularExpressions;

/***
 * CsharpExporter.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
namespace xlsx2string
{
    /// <summary>
    /// 根据表头，生成C#类定义数据结构
    /// 表头使用三行定义：字段名称、字段类型、注释
    /// </summary>
    public class CsharpExporter : ExporterBase
    {
        public String ClassComment
        {
            get;
            set;
        }

        public override void Process(bool isOver = true)
        {
            if (FieldList == null) {
                throw new Exception("Filed csharp is null.");
            }

            string defName = GetFileName(Option.CSharpPath);

            // 创建代码字符串
            StringBuilder sbTab = new StringBuilder();
            sbTab.AppendLine(
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

            sbTab.AppendLine();

            sbTab.AppendLine("using xsj.framework;");
            sbTab.AppendLine();

            sbTab.AppendLine("namespace xsj");
            sbTab.AppendLine("{");

            if (this.ClassComment != null) {
                sbTab.AppendLine(this.ClassComment);
            }
            sbTab.AppendFormat("\tpublic class {0}  : AbsTabData\r\n\t{{", defName.Replace("_", ""));
            sbTab.AppendLine();

            Dictionary<string, List<FieldDef>> repeatFieldmap = new Dictionary<string, List<FieldDef>>();
            foreach (FieldDef field in FieldList) {
                sbTab.AppendFormat("\t\t// {0}", field.comment);
                sbTab.AppendLine();
                string fieldType = field.type.ToLower().Trim();
                string fieldTypeMapping = null;
                if (!TypeArray.TryGetValue(fieldType, out fieldTypeMapping)) {
                    throw new Exception(string.Format("Error type {0}", fieldType));
                }
                sbTab.AppendFormat("\t\tpublic {0} {1}", fieldTypeMapping, field.name);
                sbTab.AppendLine();
                sbTab.AppendLine("\t\t{");
                sbTab.AppendFormat("\t\t\tget;");
                sbTab.AppendLine();
                sbTab.AppendFormat("\t\t\tset;");
                sbTab.AppendLine();
                sbTab.AppendLine("\t\t}");

                string suffix = field.name.Trim().Substring(field.name.Length - 1);
                string prefix = field.name.Trim().Substring(0, field.name.Length - 1);
               
                if (IsInt(suffix)) {
                    List<FieldDef> fieldList = null;
                    if (!repeatFieldmap.TryGetValue(prefix, out fieldList)) {
                        fieldList = new List<FieldDef>();
                        repeatFieldmap.Add(prefix, fieldList);
                    }
                    fieldList.Add(field);
                }
            }

            foreach(KeyValuePair<string, List<FieldDef>> pair in repeatFieldmap) {
                sbTab.AppendLine();
                sbTab.AppendFormat("\t\tpublic {0}[] Get{1}Array() \n\t\t{{\n", TypeArray[pair.Value[0].type.ToLower().Trim()], pair.Key.Replace("_", string.Empty));
                sbTab.AppendFormat("\t\t\treturn new {0}[] {{", TypeArray[pair.Value[0].type.ToLower().Trim()]);
                int index = 0;
                foreach (FieldDef field in pair.Value) {
                    index++;
                    if (index == pair.Value.Count) {
                        sbTab.AppendFormat("{0}", field.name);
                    } else {
                        sbTab.AppendFormat("{0}, ", field.name);
                    }
                }
                sbTab.Append(" };");
                sbTab.AppendLine();
                sbTab.AppendLine("\t\t}");
            }

            sbTab.Append("\t}");

            sbTab.AppendLine();
            sbTab.AppendLine();


            // 创建代码字符串
            StringBuilder sbConf = new StringBuilder();
            sbConf.AppendFormat("\tpublic class {0}Config : AbsTabConfig\r\n\t{{", defName.Replace("_", ""));
            sbConf.AppendLine();
            sbConf.AppendFormat("\t\tprivate const string FILE_NAME = \"Assets/Arts/Config/Tab/{0}.tab.txt\";\n", defName.ToLower().Replace("_", ""));
            sbConf.AppendLine();
            sbConf.AppendLine("\t\tpublic enum Cols");
            sbConf.AppendLine("\t\t{");
            foreach (FieldDef field in FieldList) {
                sbConf.AppendFormat("\t\t\t{0},\n", field.name.ToUpper());
            }
            sbConf.AppendLine("\t\t}");
            sbConf.AppendLine();
            sbConf.AppendLine("\t\tpublic override void Init()");
            sbConf.AppendLine("\t\t{");
            sbConf.AppendLine("\t\t\tConfig.Load<TabReaderImpl>(FILE_NAME, this, false);");
            sbConf.AppendLine("\t\t}");
            sbConf.AppendLine();

            sbConf.AppendLine("\t\tpublic override void OnRow(ITabRow row) \n\t\t{");
            sbConf.AppendFormat("\t\t\t{0} tab = new {1}();\n", defName.Replace("_", ""), defName.Replace("_", ""));
            foreach (FieldDef field in FieldList) {
                string typeName = field.type.Substring(0, 1).ToUpper() + field.type.Substring(1);
                sbConf.AppendFormat("\t\t\ttab.{0} = row.Get{1}((int)Cols.{2});\n", field.name, typeName, field.name.ToUpper());
            }
            sbConf.AppendLine();
            sbConf.AppendFormat("\t\t\tif (!ConfigMap.ContainsKey(tab.{0}.ToString())) {{\n", FieldList[0].name);
            sbConf.AppendFormat("\t\t\t\tConfigMap.Add(tab.{0}.ToString(), tab);\n", FieldList[0].name);
            sbConf.AppendLine("\t\t\t}");
            sbConf.AppendLine("\t\t}");
            sbConf.AppendLine("\t}");
            sbConf.AppendLine("}");

            StringBuilder sb = new StringBuilder();
            sb.Append(sbTab.ToString());
            sb.Append(sbConf.ToString());

            // 写文件
            string p = GetFileName(Option.CSharpPath);
            p = string.Format("{0}/{1}Config.cs", "cs", p.Replace("_", ""));
            p = Path.Combine(Path.GetDirectoryName(Option.CSharpPath), p);

            // 写文件
            WriteFile(p, sb.ToString(), Coding);
        }

        public bool IsInt(string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*$");
        }
    }
}
