using System;
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
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("//");
            sb.AppendLine("// Auto Generated Code By Text");
            sb.AppendLine("//");
            sb.AppendLine();
            if (this.ClassComment != null) {
                sb.AppendLine(this.ClassComment);
            }
            sb.AppendFormat("public class {0}\r\n{{", defName);
            sb.AppendLine();

            foreach (FieldDef field in FieldList) {
                sb.AppendFormat("\t// {0}", field.comment);
                sb.AppendLine();
                sb.AppendFormat("\tpublic {0} {1}", field.type, field.name);
                sb.AppendLine(" {");
                sb.AppendFormat("\t\tget;");
                sb.AppendLine();
                sb.AppendFormat("\t\tset;");
                sb.AppendLine();
                sb.AppendLine("\t}");
            }

            sb.Append('}');
            sb.AppendLine();
            sb.AppendLine("// End of Auto Generated Code");

            // 写文件
            WriteFile(Option.CSharpPath, sb.ToString(), Coding);
        }
    }
}
