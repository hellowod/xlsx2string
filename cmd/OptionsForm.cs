using System.Collections.Generic;

namespace xlsx2string
{
    /// <summary>
    /// OptionsForm 选项参数
    /// </summary>
    public class OptionsForm
    {
        public string XlsxSrcPath
        {
            get;
            set;
        }

        public string XlsxDstPath
        {
            get;
            set;
        }

        public List<ExportType> ExporterList
        {
            get;
            private set;
        }

        public OptionsForm()
        {
            this.XlsxSrcPath = "";
            this.XlsxDstPath = "";

            this.ExporterList = new List<ExportType>();
        }

        public void SetExportType(ExportType type)
        {
            if (ExporterList.Contains(type)) {
                return;
            }
            this.ExporterList.Add(type);
        }

        public void RemExportType(ExportType type)
        {
            if (!ExporterList.Contains(type)) {
                return;
            }
            this.ExporterList.Remove(type);
        }
    }
}
