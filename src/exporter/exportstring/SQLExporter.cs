using System.Data;
using System.Text;

/***
 * SQLExporter.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
namespace xlsx2string
{
    public class SQLExporter : ExporterBase
    {
        private int m_headerRows = 2;

        /// <summary>
        /// 将表单内容转换成INSERT语句
        /// </summary>
        private string GetTableContentSQL(DataTable sheet, string tabelName)
        {
            StringBuilder sbContent = new StringBuilder();
            StringBuilder sbNames = new StringBuilder();
            StringBuilder sbValues = new StringBuilder();

            // 字段名称列表
            foreach (DataColumn column in sheet.Columns) {
                sbNames.Append(column.ToString());
                sbNames.Append(", ");
            }

            // 逐行转换数据
            int firstDataRow = m_headerRows - 1;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++) {
                DataRow row = sheet.Rows[i];
                sbValues.Clear();
                foreach (DataColumn column in sheet.Columns) {
                    if (sbValues.Length > 0) {
                        sbValues.Append(", ");
                    }
                    sbValues.AppendFormat("'{0}'", row[column].ToString());
                }

#if false
                sbContent.AppendFormat("INSERT INTO `{0}`({1}) VALUES({2});\n",
                    tabelName, sbNames.ToString(), sbValues.ToString());
#else
                sbContent.AppendFormat("INSERT INTO `{0}` VALUES({1});\n",
                    tabelName, sbValues.ToString());
#endif

            }
            return sbContent.ToString();
        }

        /// <summary>
        /// 根据表头构造CREATE TABLE语句
        /// </summary>
        private string GetTabelStructSQL(DataTable sheet, string tabelName)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("DROP TABLE IF EXISTS `{0}`;\n", tabelName);
            sb.AppendFormat("CREATE TABLE `{0}` (\n", tabelName);

            DataRow typeRow = sheet.Rows[0];

            foreach (DataColumn column in sheet.Columns) {
                string filedName = column.ToString();
                string filedType = typeRow[column].ToString();
                if (filedType == "varchar") {
                    sb.AppendFormat("`{0}` {1}(64),", filedName, filedType);
                } else if (filedType == "text") {
                    sb.AppendFormat("`{0}` {1}(256),", filedName, filedType);
                } else {
                    sb.AppendFormat("`{0}` {1},", filedName, filedType);
                }
            }
            sb.AppendFormat("PRIMARY KEY (`{0}`) ", sheet.Columns[0].ToString());
            sb.AppendLine("\n) DEFAULT CHARSET=utf8;");
            return sb.ToString();
        }

        public override void Init()
        {
        }

        public override void Export()
        {
            string tableName = GetFileName(Option.JsonPath);
            string tabelStruct = GetTabelStructSQL(Sheet, tableName);
            string tabelContent = GetTableContentSQL(Sheet, tableName);

            StringBuilder sb = new StringBuilder();
            sb.Append(tabelStruct);
            sb.AppendLine();
            sb.Append(tabelContent);

            WriteFile(Option.SQLPath, sb.ToString(), Coding);
        }
    }
}
