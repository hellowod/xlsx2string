using System;

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
    }
}
