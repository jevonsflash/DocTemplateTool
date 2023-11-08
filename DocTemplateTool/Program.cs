
using System.Diagnostics;
using DocTemplateTool.Common.Helper;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.XWPF.UserModel;

namespace DocTemplateTool
{
    partial class Program
    {
        static void Main(string[] args)
        {
            if (!CliProcessor.ProcessCommandLine(args))
            {
                Console.WriteLine("缺少参数或参数不正确");

                CliProcessor.Usage();
                Environment.ExitCode = 1;
                return;
            }
            try
            {

                var sw = Stopwatch.StartNew();
                IDictionary<string, object> docinfo = null;
                if (CliProcessor.source == "json")
                {
                    var docinfoJson = DirFileHelper.ReadFile(CliProcessor.objectFilePath);
                    docinfo = CommonHelper.ToCollections(JObject.Parse(docinfoJson)) as IDictionary<string, object>;
                }
                else
                {
                }

                if (CliProcessor.destination == "word")
                {
                    var output = WordDocProcessor.ImportFrom(CliProcessor.inputPathList.First(), docinfo);
                    WordDocProcessor.SaveTo(output, CliProcessor.outputPathList.First());
                    Console.WriteLine("已成功完成");
                }
                else if (CliProcessor.destination == "pdf")
                {

                    var output = PdfDocProcessor.ImportFrom(CliProcessor.inputPathList.First(), docinfo);
                    PdfDocProcessor.SaveTo(output, CliProcessor.outputPathList.First());
                    Console.WriteLine("已成功完成");
                }

                sw.Stop();


                Console.WriteLine("Time taken: {0}", sw.Elapsed);
            }
            catch (Exception ex)
            {
                Console.WriteLine("{0}未知错误:{0}{1}", Environment.NewLine, ex);
                Environment.ExitCode = 2;
            }

            if (CliProcessor.waitAtEnd)
            {
                Console.WriteLine("{0}{0}敲击回车退出程序", Environment.NewLine);
                Console.ReadLine();
            }
        }


    }
}
