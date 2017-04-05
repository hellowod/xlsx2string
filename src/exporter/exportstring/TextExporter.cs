﻿using System;
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
            foreach (DataRow row in Sheet.Rows) {
                object[] columns = row.ItemArray;
                sb.Append(columns[0]);
                foreach (object obj in columns) {
                    sb.AppendFormat("\t{0}", obj);
                }
                sb.AppendLine();
            }
            // 写文件
            this.WriteFile(Option.TxtPath, sb.ToString(), Coding);
        }

        public override void Init()
        {
            
        }

        public override void SaveToFile(string filePath, Encoding encoding)
        {
            throw new NotImplementedException();
        }
    }
}
