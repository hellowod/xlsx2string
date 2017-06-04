using System.Data;
using System.Text;

/***
 * LuaExporter.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.07
 */
namespace xlsx2string
{
    public class LuaExporter : ExporterBase
    {
        private int m_headerRows = 3;

        private string GetContextLua()
        {
            StringBuilder sbContent = new StringBuilder();
            StringBuilder sbValues = new StringBuilder();

            sbContent.AppendFormat("{0} = ", GetFileName(Option.LuaPath));
            sbContent.Append("{\n");

            // 逐行转换数据
            int firstDataRow = m_headerRows - 1;
            for (int i = firstDataRow; i < Sheet.Rows.Count; i++) {
                DataRow row = Sheet.Rows[i];
                sbValues.Clear();
                sbValues.AppendFormat("[{0}]=[", row[Sheet.Columns[0]].ToString());
                for(int j = 1; j < Sheet.Columns.Count; ++j) {
                    if(j > 1) {
                        sbValues.Append(", ");
                    }
                    sbValues.AppendFormat("{0}", row[Sheet.Columns[j]].ToString());
                }
                sbValues.Append("],");

                sbContent.AppendFormat("{0}\n", sbValues.ToString());
            }
            sbContent.Append("}");

            return sbContent.ToString();
        }

        private string GetStructLua()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("local function convertFunc(key, arr)");
            sb.AppendLine();
            sb.Append("\treturn {");
            DataRow typeRow = Sheet.Rows[0];
            for (int i = 0; i < Sheet.Columns.Count; ++i) {
                string filedName = Sheet.Columns[i].ToString();
                if (i == 0) {
                    sb.AppendFormat("{0}={1} ", filedName, "key");
                } else {
                    sb.AppendFormat("{0}=arr[{1}] ", filedName, i);
                }
            }
            sb.Append(",}");
            sb.AppendLine();
            sb.Append("end");
            sb.AppendLine();

            string fileName = GetFileName(Option.LuaPath);

            sb.AppendFormat("{0} = ", fileName);
            sb.Append("{");
            sb.AppendFormat("_src_tb_ = _{0}, _tb_convert_func_=convertFunc", fileName);
            sb.Append(" }");
            sb.AppendLine();
            sb.AppendFormat("setmetatable({0}, _tb_cfg_mt_)", fileName);
            sb.AppendLine();

            return sb.ToString();
        }

        public override void Process()
        {
            string tabelStruct = GetStructLua();
            string tabelContent = GetContextLua();

            StringBuilder sb = new StringBuilder();
            sb.Append(tabelStruct);
            sb.AppendLine();
            sb.Append(tabelContent);

            WriteFile(Option.LuaPath, sb.ToString(), Coding);
        }

        public override void Init()
        {
            
        }
    }
}
