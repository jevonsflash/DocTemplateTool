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
using PdfSharpCore.Pdf;
using DocTemplateTool.Pdf;

namespace DocTemplateTool
{
    public class PdfDocProcessor
    {

        public static void SaveTo(PdfDocument result, string filePath)
        {

            var extension = Path.GetExtension(filePath).ToLower();

            if (extension.EndsWith(".pdf"))
            {

                try
                {
                    result.Save(filePath);

                }
                catch (Exception e)
                {
                    Console.WriteLine(filePath + " 此文件正被其他程序占用");
                }

            }
            else
            {
                Console.WriteLine(filePath + " 文件格式不合法");
            }

        }




        public static PdfDocument ImportFrom(string filePath, IDictionary<string, object> docinfo)
        {

            PdfDocument output = null;

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
                if (filePath.EndsWith(".pdf"))
                {
                    using (var stream = new MemoryStream(data1))
                    {
                        output = Exporter.ExportDocxByDictionary(stream, docinfo, (s) => s);
                    }


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
