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
    public class CsvExporter : ExporterBase
    {
        public override void Process()
        {
            StringBuilder sb = new StringBuilder();
            foreach (DataRow row in Sheet.Rows) {
                object[] columns = row.ItemArray;
                sb.Append(columns[0]);
                foreach (object obj in columns) {
                    sb.AppendFormat("{0},", obj);
                }
                sb.AppendLine();
            }
            // 写文件
            this.WriteFile(Option.CsvPath, sb.ToString(), Coding);
        }

        public override void Init()
        {
            
        }
    }
}
