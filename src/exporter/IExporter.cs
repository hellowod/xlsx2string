using System.Data;
using System.Text;

namespace xlsx2string
{
    public interface IExporter
    {
        DataTable Sheet
        {
            get;
            set;
        }

        Options Option
        {
            get;
            set;
        }

        Encoding Coding
        {
            get;
            set;
        }

        void Init();

        void Export();

        void SaveToFile(string filePath, Encoding encoding);
    }
}
