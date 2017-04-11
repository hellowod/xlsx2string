using System.IO;
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
    public class TextExporter : ExporterBase
    {
        public override void Export()
        {
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < ColCount; ++i) {
                FieldDef field = FieldList[i];
                if(i >= (ColCount - 1)) {
                    sb.AppendFormat("{0}", field.name);
                } else {
                    sb.AppendFormat("{0}\t", field.name);
                }
            }
            sb.AppendLine();

            foreach (DataRow row in Sheet.Rows) {
                object[] columns = row.ItemArray;
                sb.Append(columns[0]);
                foreach (object obj in columns) {
                    sb.AppendFormat("\t{0}", obj);
                }
                sb.AppendLine();
            }

            // 写文件
            string p = GetFileName(Option.TxtPath);
            p = string.Format("{0}.tab.txt", p.ToLower());
            p = Path.Combine(Path.GetDirectoryName(Option.TxtPath), p);

            this.WriteFile(p, sb.ToString(), Coding);
        }
    }
}
