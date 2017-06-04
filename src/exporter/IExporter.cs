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

        /// <summary>
        /// 实例化
        /// </summary>
        void New();

        /// <summary>
        /// 初始化
        /// </summary>
        void Init();

        /// <summary>
        /// 处理
        /// </summary>
        void Process();

        /// <summary>
        /// 清理
        /// </summary>
        void Clear();
    }
}
