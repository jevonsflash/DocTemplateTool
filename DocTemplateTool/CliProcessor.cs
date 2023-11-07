using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DocTemplateTool
{
    partial class CliProcessor
    {
        public static List<string> inputPathList;
        public static List<string> outputPathList;
        public static string destination;
        public static string source;
        public static bool waitAtEnd;
        public static string objectFilePath;

        public static void Usage()
        {
            var versionInfo = FileVersionInfo.GetVersionInfo(Environment.ProcessPath);
            Console.WriteLine();
            Console.WriteLine("Document Template Tool v{0}.{1}", versionInfo.FileMajorPart, versionInfo.FileMinorPart);
            Console.WriteLine("参数列表:");
            Console.WriteLine(" -p  ObjectFile");
            Console.WriteLine("     指定一个Object文件(Json), 作为数据源");
            Console.WriteLine(" -i  Input");
            Console.WriteLine("     指定一个docx文件作为模板");
            Console.WriteLine(" -o  Output");
            Console.WriteLine("     指定一个路径，作为导出目标");
            Console.WriteLine(" -s  Source");
            Console.WriteLine("     值为json");
            Console.WriteLine(" -d  Destination");
            Console.WriteLine("     值为word, pdf");
            Console.WriteLine(" -w  WaitAtEnd");
            Console.WriteLine("     指定时，程序执行完成后，将等待用户输入退出");
            Console.WriteLine(" -h  Help");
            Console.WriteLine("     查看帮助");
        }


        public static bool ProcessCommandLine(string[] args)
        {
            inputPathList = new List<string>();
            outputPathList = new List<string>();
            destination = string.Empty;
            source = string.Empty;
            var i = 0;
            while (i < args.Length)
            {
                var arg = args[i];
                if (arg.StartsWith("/") || arg.StartsWith("-"))
                    arg = arg.Substring(1);
                switch (arg.ToLowerInvariant())
                {

                    case "p":
                        i++;
                        if (i < args.Length)
                        {
                            if (!File.Exists(args[i]))
                            {
                                Console.WriteLine("文件 '{0}' 不存在", args[i]);
                                return false;
                            }
                            objectFilePath = args[i];
                        }
                        else
                            return false;
                        break;
                    case "i":
                        i++;
                        if (i < args.Length)
                        {
                            if (!File.Exists(args[i]))
                            {
                                Console.WriteLine("文件 '{0}' 不存在", args[i]);
                                return false;
                            }
                            inputPathList.Add(args[i]);
                        }
                        else
                            return false;
                        break;
                    case "s":
                        i++;
                        if (i < args.Length)
                        {

                            if (!new string[] { "json" }.Any(c => c == args[i]))
                            {
                                Console.WriteLine("参数值 '{0}' 不合法", args[i]);
                                return false;
                            }
                            source = args[i];

                        }
                        else
                            return false;
                        break;
                    case "o":
                        i++;
                        if (i < args.Length)
                        {
                            if (args[i].IndexOfAny(Path.GetInvalidPathChars()) >= 0)
                            {
                                Console.WriteLine("路径 '{0}' 不合法", args[i]);
                                return false;
                            }
                            outputPathList.Add(args[i]);
                        }
                        else
                            return false;
                        break;
                    case "d":
                        i++;
                        if (i < args.Length)
                        {
                            if (!new string[] { "word", "pdf" }.Any(c => c == args[i]))
                            {
                                Console.WriteLine("参数值 '{0}' 不合法", args[i]);
                                return false;
                            }
                            destination = args[i];
                        }
                        else
                            return false;
                        break;
                    case "w":
                        waitAtEnd = true;
                        break;
                    case "h":
                        Usage();
                        return false;


                    default:
                        Console.WriteLine("无法识别的参数: {0}", args[i]);
                        break;
                }
                i++;
            }
            return !string.IsNullOrEmpty(objectFilePath)
                && !string.IsNullOrEmpty(source)
                && !string.IsNullOrEmpty(destination)
                && inputPathList.Count > 0
                && outputPathList.Count > 0;

        }
    }
}