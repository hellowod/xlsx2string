using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsx2string
{
    public interface IExporter
    {
        void SaveToFile(string filePath, Encoding encoding);

        void ToFile(DataTable sheet, Options option, Encoding encoding);
    }
}
