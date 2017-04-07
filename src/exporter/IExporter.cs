using System.Data;
using System.Text;

/***
 * IExporter.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
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

        void Clear();
    }
}
