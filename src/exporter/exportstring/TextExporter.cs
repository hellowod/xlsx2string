using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsx2string
{
    public class TextExporter : IExporter
    {
        public Encoding Coding
        {
            get;
            set;
        }

        public Options Option
        {
            get;
            set;
        }

        public DataTable Sheet
        {
            get;
            set;
        }

        public void Export()
        {
            throw new NotImplementedException();
        }

        public void Init()
        {
            throw new NotImplementedException();
        }

        public void SaveToFile(string filePath, Encoding encoding)
        {
            throw new NotImplementedException();
        }
    }
}
