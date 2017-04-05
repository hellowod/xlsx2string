using System.Data;
using System.IO;
using System.Text;

/***
 * ExporterBase.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
namespace xlsx2string
{
    public abstract class ExporterBase : IExporter
    {
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

        protected virtual void WriteFile(string path, string context, Encoding coding)
        {
            string outPath = Path.GetDirectoryName(path);
            if (!Directory.Exists(outPath)) {
                Directory.CreateDirectory(outPath);
            }
            if (File.Exists(path)) {
                File.Delete(path);
            }

            using (FileStream file = new FileStream(path, FileMode.Create, FileAccess.Write)) {
                using (TextWriter writer = new StreamWriter(file, coding)) {
                    writer.Write(context);
                }
            }
        }

        protected virtual string GetFileName(string path)
        {
            return Path.GetFileNameWithoutExtension(path);
        }

        public virtual void Init()
        {
            
        }

        public virtual void Export()
        {
            
        }

        public virtual void SaveToFile(string filePath, Encoding encoding)
        {
            
        }
    }
}
