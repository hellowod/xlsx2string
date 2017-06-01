﻿using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;

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

        public override void Export()
        {
            if (FieldList == null) {
                throw new Exception("Filed csharp is null.");
            }

            string defName = GetFileName(Option.CSharpPath);

            // 创建代码字符串
            StringBuilder sbTab = new StringBuilder();
            sbTab.AppendLine("//");
            sbTab.AppendLine("// Auto Generated Code By Text");
            sbTab.AppendLine("// @Author abaojin");
            sbTab.AppendLine("//");
            sbTab.AppendLine();

            sbTab.AppendLine("using GameEngine;");
            sbTab.AppendLine();

            if (this.ClassComment != null) {
                sbTab.AppendLine(this.ClassComment);
            }
            sbTab.AppendFormat("public class {0}Tab  : IConfData\r\n{{", defName);
            sbTab.AppendLine();

            foreach (FieldDef field in FieldList) {
                sbTab.AppendFormat("\t// {0}", field.comment);
                sbTab.AppendLine();
                sbTab.AppendFormat("\tpublic {0} {1}", field.type.ToLower(), field.name);
                sbTab.AppendLine(" {");
                sbTab.AppendFormat("\t\tget;");
                sbTab.AppendLine();
                sbTab.AppendFormat("\t\tset;");
                sbTab.AppendLine();
                sbTab.AppendLine("\t}");
            }

            sbTab.Append('}');

            sbTab.AppendLine();
            sbTab.AppendLine();


            // 创建代码字符串
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("public class {0}TabConf : AbsTabConf\r\n{{", defName);
            sb.AppendLine();
            sb.AppendFormat("\tpublic const string FILE_NAME = \"{0}.tab\";\n", defName.ToLower());
            sb.AppendLine();
            sb.AppendLine("\tpublic enum Cols");
            sb.AppendLine("\t{");
            foreach (FieldDef field in FieldList) {
                sb.AppendFormat("\t\t{0},\n", field.name.ToUpper());
            }
            sb.AppendLine("\t}");
            sb.AppendLine();
            sb.AppendLine("\tpublic override void Init()");
            sb.AppendLine("\t{");
            sb.AppendLine("\t\tConfFactory.LoadConf<TabReaderImpl>(FILE_NAME, this);");
            sb.AppendLine("\t}");
            sb.AppendLine();

            sb.AppendLine("\tpublic override void OnRow(ITabRow row) {");
            sb.AppendFormat("\t\t{0}Tab tab = new {1}Tab();\n", defName, defName);
            foreach (FieldDef field in FieldList) {
                string typeName = field.type.Substring(0, 1).ToUpper() + field.type.Substring(1);
                sb.AppendFormat("\t\ttab.{0} = row.Get{1}((int)Cols.{2});\n", field.name, typeName, field.name.ToUpper());
            }
            sb.AppendLine();
            sb.AppendFormat("\t\tif (!ConfPool.ContainsKey(tab.{0}.ToString())) {{\n", FieldList[0].name);
            sb.AppendFormat("\t\t\tConfPool.Add(tab.{0}.ToString(), tab);\n", FieldList[0].name);
            sb.AppendLine("\t\t}");
            sb.AppendLine("\t}");
            sb.AppendLine("}");

            sb.AppendLine("// End of Auto Generated Code");

            StringBuilder sb = new StringBuilder();
            sb.Append(sbTab.ToString());
            sb.Append(sb.ToString());

            // 写文件
            string p = GetFileName(Option.CSharpPath);
            p = string.Format("{0}TabConf.cs", p);
            p = Path.Combine(Path.GetDirectoryName(Option.CSharpPath), p);

            // 写文件
            WriteFile(p, sb.ToString(), Coding);
        }
    }
}
