﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsx2string
{
    public class LuaExporter : IExporter
    {
        public void SaveToFile(string filePath, Encoding encoding)
        {
            throw new NotImplementedException();
        }

        public void ToFile(DataTable sheet, Options option, Encoding encoding)
        {
            throw new NotImplementedException();
        }
    }
}
