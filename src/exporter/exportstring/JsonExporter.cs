using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;

/***
 * JsonExporter.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
namespace xlsx2string
{
    /// <summary>
    /// 将DataTable对象，转换成JSON string，并保存到文件中
    /// </summary>
    public class JsonExporter : ExporterBase
    {
        private Dictionary<string, Dictionary<string, object>> datTable;

        public override void Process()
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
            base.Init();

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
    }
}
