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
        private static Dictionary<string, string> arrayType;

        private List<FieldDef> fieldList;

        static ExporterBase()
        {
            arrayType = new Dictionary<string, string>();
        }

        public ExporterBase()
        {
            this.New();
        }

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

        public Dictionary<string, string> TypeArray
        {
            get {
                return arrayType;
            } 
        }

        public List<FieldDef> FieldList
        {
            get {
                return fieldList;
            }
        }

        /// <summary>
        /// 默认映射关系（可重写）
        /// </summary>
        protected virtual void TypeMapping()
        {
            TypeArray["int"] = "int";
            TypeArray["uint"] = "uint";
            TypeArray["short"] = "short";
            TypeArray["ushort"] = "ushort";
            TypeArray["long"] = "long";
            TypeArray["ulong"] = "ulong";
            TypeArray["string"] = "string";
            TypeArray["char"] = "char";
            TypeArray["float"] = "float";
            TypeArray["double"] = "double";
            TypeArray["bool"] = "bool";
            TypeArray["boolean"] = "bool";

            TypeArray["shortarray"] = "short[]";
            TypeArray["ushortarray"] = "ushort[]";
            TypeArray["intarray"] = "int[]";
            TypeArray["uintarray"] = "uint[]";
            TypeArray["longarray"] = "long[]";
            TypeArray["ulongarray"] = "ulong[]";
            TypeArray["stringarray"] = "string[]";
            TypeArray["floatarray"] = "float[]";
            TypeArray["doublearray"] = "double[]";
            TypeArray["boolarray"] = "bool[]";
        }

        /// <summary>
        /// 写文件
        /// </summary>
        /// <param name="path"></param>
        /// <param name="context"></param>
        /// <param name="coding"></param>
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

        /// <summary>
        /// 获取文件名称
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
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

        /// <summary>
        /// 导出器实例化
        /// </summary>
        public virtual void New()
        {
            this.TypeMapping();
        }

        /// <summary>
        /// 导出初始
        /// </summary>
        public virtual void Init()
        {
            this.ParseFiledList();
        }

        /// <summary>
        /// 处理
        /// </summary>
        public virtual void Process(bool isOver = true)
        {
            
        }

        /// <summary>
        /// 导出清理
        /// </summary>
        public virtual void Clear()
        {
           
        }
    }
}
