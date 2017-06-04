using System.Data;
using System.Text;

/***
 * TextExporter.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
namespace xlsx2string
{
    public class HtmlExporter : ExporterBase
    {
        public override void Process()
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("<!doctype html>");
            sb.AppendLine("<html lang=\"en\">");
            sb.AppendLine(" <head>");
            sb.AppendLine("  <meta charset=\"UTF - 8\">");
            sb.AppendLine("  <meta name=\"Generator\" content=\"EditPlus®\">");
            sb.AppendLine("  <meta name=\"Author\" content=\"\">");
            sb.AppendLine("  <meta name=\"Keywords\" content=\"\">");
            sb.AppendLine("  <meta name=\"Description\" content=\"\">");
            sb.AppendLine("  <title>Document</title>");
            sb.AppendLine(" </head>");
            sb.AppendLine(" <body>");

            sb.AppendLine("\t<table border=\"1\">");
            foreach (DataRow row in Sheet.Rows) {
                object[] columns = row.ItemArray;
                sb.AppendLine("\t\t<tr>");
                foreach (object obj in columns) {
                    sb.AppendFormat("\t\t\t<th>{0}</th>\n", obj);
                }
                sb.Append("\t\t</tr>");
                sb.AppendLine();
            }
            sb.AppendLine("\t</table>");
            sb.AppendLine(" </body>");
            sb.AppendLine("</html>");

            // 写文件
            this.WriteFile(Option.HtmlPath, sb.ToString(), Coding);
        }

        public override void Init()
        {
            
        }
    }
}
