using System;
using System.Windows.Forms;
using System.IO;

/***
 * PathUtils.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
namespace xlsx2string
{
    public static class PathUtils
    {
        /// <summary>
        /// 获取应用数据目录
        /// </summary>
        /// <returns></returns>
        public static string GetLocalApplicationDataPath()
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        }

        /// <summary>
        /// 产品的本地数据目录
        /// </summary>
        /// <returns></returns>
        public static string GetProductLocalApplicationDataPath()
        {
            return Path.Combine(GetLocalApplicationDataPath(), Application.ProductName);
        }

        /// <summary>
        /// 产品缓存文件路径
        /// </summary>
        /// <returns></returns>
        public static string GetProductCacheFilePath()
        {
            return Path.Combine(GetProductLocalApplicationDataPath(), string.Format("{0}.json", Application.ProductName));
        }
    }
}
