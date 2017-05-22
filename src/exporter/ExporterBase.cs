using System.Data;
using System.IO;
using System.Text;
using System.Collections.Generic;

/***
 * ExporterBase.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
namespace xlsx2string
{
    public struct FieldDef
    {
        public string name;
        public string type;
        public string comment;
    }

    public abstract class ExporterBase : IExporter
    {
        private List<FieldDef> fieldList;

        public DataTable Sheet
        {
            get;
            set;
        }

        public Options Option
        {
            get;
            set;
        }

        public Encoding Coding
        {
            get;
            set;
        }

        public int ColCount
        {
            get;
            private set;
        }

        public List<FieldDef> FieldList
        {
            get {
                return fieldList;
            }
        }

        protected virtual void WriteFile(string path, string context, Encoding coding)
        {
            string outPath = Path.GetDirectoryName(path);
            if (!Directory.Exists(outPath)) {
                Directory.CreateDirectory(outPath);
            }
            if (File.Exists(path)) {
                File.Delete(path);
            }

            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write)) {
                using (TextWriter writer = new StreamWriter(stream, coding)) {
                    writer.Write(context);
                }
            }
        }

        protected virtual string GetFileName(string path)
        {
            return Path.GetFileNameWithoutExtension(path);
        }

        /// <summary>
        /// 处理表前三行作为Filed
        /// </summary>
        protected virtual void ParseFiledList()
        {
            if (Sheet.Rows.Count < 2) {
                return;
            }

            fieldList = new List<FieldDef>();
            DataRow typeRow = Sheet.Rows[0];
            DataRow commRow = Sheet.Rows[1];

            foreach (DataColumn column in Sheet.Columns) {
                FieldDef field;
                field.name = column.ToString();
                field.type = typeRow[column].ToString();
                field.comment = commRow[column].ToString();

                fieldList.Add(field);
            }

            ColCount = fieldList.Count;
        }

        public virtual void Init()
        {
            ParseFiledList();
        }

        public virtual void Export()
        {
            
        }

        public void Clear()
        {
           
        }
    }
}
