using System;

namespace xlsx2string
{
    public enum ExportType
    {
        cs,
        txt,
        csv,
        lua,

        java,
        cpp,
        go,
        sql
    }

    public static class TypeHelper
    {
        /// <summary>转换为枚举(该函数效率低下)</summary>
        /// <typeparam name="T">枚举类型</typeparam>
        /// <param name="val">值</param>
        /// <param name="ignoreCase">是否忽略大小写</param>
        /// <returns></returns>
        public static T ToEnum<T>(this int val, bool ignoreCase = true)
        {
            return (T)Enum.Parse(typeof(T), val.ToString(), ignoreCase);
        }

        /// <summary>转换为枚举(该函数效率低下)</summary>
        /// <typeparam name="T">枚举类型</typeparam>
        /// <param name="val">值</param>
        /// <param name="ignoreCase">是否忽略大小写</param>
        /// <returns></returns>
        public static T ToEnum<T>(this string val, bool ignoreCase = true)
        {
            return (T)Enum.Parse(typeof(T), val, ignoreCase);
        }
    }
}
