using System;
using System.Text;

/***
 * JavaExporter.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
namespace xlsx2string
{
    public class JavaExporter : ExporterBase
    {
        public override void Process(bool isOver = true)
        {
            if (FieldList == null) {
                throw new Exception("Filed java is null.");
            }

            string defName = GetFileName(Option.JavaPath);

            // 创建代码字符串
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("//");
            sb.AppendLine("// Auto Generated Code By Text");
            sb.AppendLine("//");
            sb.AppendLine();
            sb.AppendFormat("public class {0}\r\n{{", defName);
            sb.AppendLine();

            foreach (FieldDef field in FieldList) {
                sb.AppendFormat("\t// {0}", field.comment);
                sb.AppendLine();
                sb.AppendFormat("\tprivate {0} {1};\n", field.type, field.name);
                sb.AppendFormat("\tpublic {0} get{1}()", field.type, field.name);
                sb.Append("{\n");
                sb.AppendFormat("\t\treturn this.{0};\n", field.name);
                sb.Append("\t}\n");
                sb.AppendFormat("\tpublic void set{0}({1} {2})", field.name, field.type, field.name.ToLower());
                sb.Append("{\n");
                sb.AppendFormat("\t\tthis.{0} = {1}\n", field.name, field.name.ToLower());
                sb.Append("\t}\n");
                sb.AppendLine();
            }

            sb.Append('}');
            sb.AppendLine();
            sb.AppendLine("// End of Auto Generated Code");

            // 写文件
            WriteFile(Option.JavaPath, sb.ToString(), Coding);
        }

        public override void Init()
        {
            base.Init();
        }
    }
}
