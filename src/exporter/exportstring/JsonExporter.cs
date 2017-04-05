using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace xlsx2string
{
    /// <summary>
    /// 将DataTable对象，转换成JSON string，并保存到文件中
    /// </summary>
    public class JsonExporter : ExporterBase
    {
        private Dictionary<string, Dictionary<string, object>> datTable;

        public override void Export()
        {
            if (datTable == null) {
                throw new Exception("JsonExporter内部数据为空。");
            }

            // JSON字符
            string json = JsonConvert.SerializeObject(datTable, Formatting.Indented);

            // 写文件
            this.WriteFile(Option.JsonPath, json, Coding);
        }

        public override void Init()
        {
            if (Sheet.Columns.Count <= 0) {
                return;
            }

            if (Sheet.Rows.Count <= 0) {
                return;
            }

            datTable = new Dictionary<string, Dictionary<string, object>>();

            // 以第一列为ID，转换成ID->Object的字典
            int firstDataRow = 2 - 1;
            for (int i = firstDataRow; i < Sheet.Rows.Count; i++) {
                DataRow row = Sheet.Rows[i];
                string id = row[Sheet.Columns[0]].ToString();
                if (id.Length <= 0) {
                    continue;
                }
                Dictionary<string, object> rowData = new Dictionary<string, object>();
                foreach (DataColumn column in Sheet.Columns) {
                    object value = row[column];
                    // 去掉数值字段的“.0”
                    if (value.GetType() == typeof(double)) {
                        double num = (double)value;
                        if ((int)num == num) {
                            value = (int)num;
                        }
                    }
                    string fieldName = column.ToString();
                    if (!string.IsNullOrEmpty(fieldName)) {
                        rowData[fieldName] = value;
                    }
                }
                datTable[id] = rowData;
            }
        }

        /// <summary>
        /// 将内部数据转换成Json文本，并保存至文件
        /// </summary>
        /// <param name="jsonPath">输出文件路径</param>
        public override void SaveToFile(string filePath, Encoding encoding)
        {
            if (datTable == null) {
                throw new Exception("JsonExporter内部数据为空。");
            }

            // 转换为JSON字符串
            string json = JsonConvert.SerializeObject(datTable, Formatting.Indented);

            // 保存文件
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write)) {
                using (TextWriter writer = new StreamWriter(file, encoding))
                    writer.Write(json);
            }
        }
    }
}
