using System;
using System.IO;
using System.Text;

/***
 * FileUtils.cs
 * 
 * Author abaojin
 * Version 1.0
 * Date 2017.04.05
 */
namespace xlsx2string
{
    public static class FileUtils
    {
        private static readonly Encoding Encoding = Encoding.UTF8;

        public static void CreateFile(string filePath)
        {
            string dirPath = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(dirPath)) {
                Directory.CreateDirectory(dirPath);
            }
            if (!File.Exists(filePath)) {
                File.CreateText(filePath).Close();
            }
        }

        public static void WriteFile(string filePath, string content)
        {
            try {
                FileStream fs = new FileStream(filePath, FileMode.Create);
                Encoding encode = Encoding;
                byte[] data = encode.GetBytes(content);
                fs.Write(data, 0, data.Length);
                fs.Flush();
                fs.Close();
            } catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }
        }

        public static string ReadFile(string filePath)
        {
            return ReadFile(filePath, Encoding);
        }

        public static string ReadFile(string filePath, Encoding encoding)
        {
            if (!File.Exists(filePath)) {
                return string.Empty;
            }

            using (var sr = new StreamReader(filePath, encoding)) {
                return sr.ReadToEnd();
            }
        }
    }
}
