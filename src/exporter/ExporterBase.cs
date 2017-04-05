using System.IO;
using System.Text;

namespace xlsx2string
{
    public abstract class ExporterBase
    {
        protected virtual void WriteFile(string path, string context, Encoding coding)
        {
            string outPath = Path.GetDirectoryName(path);
            if (!Directory.Exists(outPath)) {
                Directory.CreateDirectory(outPath);
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
    }
}
