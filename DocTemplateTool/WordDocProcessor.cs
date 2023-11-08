using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using NPOI.XWPF.UserModel;
using NPOI.POIFS.Crypt.Dsig;
using Org.BouncyCastle.Utilities;

namespace DocTemplateTool
{
    public class WordDocProcessor
    {

        public static void SaveTo(XWPFDocument result, string filePath)
        {

            var extension = Path.GetExtension(filePath).ToLower();
            byte[] docFileContent = null;

            if (extension.EndsWith("doc"))
            {
                Console.WriteLine(filePath + " 暂不支持Word97-2003格式");

            }
            else if (extension.EndsWith("docx"))
            {
                using (var memoryStream = new MemoryStream())
                {
                    result.Write(memoryStream);
                    memoryStream.Seek(0, SeekOrigin.Begin);
                    docFileContent = memoryStream.ToArray();
                }

            }
            else
            {
                Console.WriteLine(filePath + " 文件格式不合法");

            }
            FileStream fs;
            try
            {
                fs = new FileStream(filePath, FileMode.Create);
                fs.Write(docFileContent, 0, docFileContent.Length);
                fs.Close();


            }
            catch (Exception e)
            {
                Console.WriteLine(filePath + " 此文件正被其他程序占用");
            }


        }




        public static XWPFDocument ImportFrom(string filePath, IDictionary<string, object> docinfo)
        {

            XWPFDocument output = null;

            var data1 = new byte[0];
            try
            {
                data1 = File.ReadAllBytes(filePath);
            }
            catch (Exception e)
            {
                Console.WriteLine(filePath + " 此文件正被其他程序占用");
                return output;
            }

            try
            {
                if (filePath.EndsWith(".docx"))
                {
                    using (var stream = new MemoryStream(data1))
                    {
                        output = Word.Exporter.ExportDocxByDictionary(stream, docinfo, (s) => s);
                    }


                }
                else if (filePath.EndsWith(".doc"))
                {
                    Console.WriteLine(filePath + " 暂不支持Word97-2003格式");

                }
                else
                {
                    Console.WriteLine(filePath + " 文件类型错误");
                    return output;
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.WriteLine(filePath + " 格式错误");
            }

            return output;


        }




    }
}
