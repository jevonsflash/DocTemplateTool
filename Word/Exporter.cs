using Flurl.Http;
using NPOI.XWPF.UserModel;
using System.Reflection;
using System.Text.RegularExpressions;
using Word.Helper;

namespace Word
{
    public class Exporter
    {
        const string Dot = ".";

        public static XWPFDocument ExportDocxByObject(Stream templateStream, object data, Func<string, string> func = null)
        {
            var doc = new XWPFDocument(templateStream);
            foreach (var para in doc.Paragraphs)
            {
                ReplaceKeyObjetAsync(para, data, func);
            }

            foreach (var table in doc.Tables)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.GetTableCells())
                    {
                        foreach (var para in cell.Paragraphs)
                        {
                            ReplaceKeyObjetAsync(para, data, func);
                        }
                    }
                }
            }

            return doc;
        }


        public static XWPFDocument ExportDocxByObject(string templateFilePath, object data, Func<string, string> func = null)
        {
            using (var fileStream = new FileStream(templateFilePath, FileMode.Open))
            {
                return ExportDocxByObject(fileStream, data, func);
            }
        }


        public static XWPFDocument ExportDocxByDictionary(Stream templateStream, Dictionary<string, string> data, Func<string, string> func = null)
        {
            var doc = new XWPFDocument(templateStream);
            foreach (var para in doc.Paragraphs)
            {
                ReplaceKeyObjetAsync(para, data, func);
            }

            foreach (var table in doc.Tables)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.GetTableCells())
                    {
                        foreach (var para in cell.Paragraphs)
                        {
                            ReplaceKeyDictionaryAsync(para, data, func);
                        }
                    }
                }
            }

            return doc;
        }


        public static XWPFDocument ExportDocxByDictionary(string templateFilePath, Dictionary<string, string> data, Func<string, string> func = null)
        {
            using (var fileStream = new FileStream(templateFilePath, FileMode.Open))
            {
                return ExportDocxByDictionary(fileStream, data, func);
            }
        }



        private static async Task ReplaceKeyDictionaryAsync(XWPFParagraph para, Dictionary<string, string> data, Func<string, string> func)
        {
            string text = "";


            foreach (var run in para.Runs)
            {
                text = run.ToString();
                foreach (var p in data.Keys)
                {
                    string key = $"${p}$";
                    if (text.Contains(key))
                    {
                        try
                        {
                            text = text.Replace(key, data[key]);
                        }
                        catch (Exception ex)
                        {
                            text = text.Replace(key, "");
                        }
                    }
                    else if (text.Contains($@"#{p}#"))
                    {
                        text = text.Replace($@"#{p}#", "");
                        try
                        {
                            var filePath = data[key];
                            filePath = func?.Invoke(filePath);

                            if (string.IsNullOrEmpty(filePath))
                            {
                                continue;
                            }

                            if (File.Exists(filePath))
                            {
                                using (var fileStream = new FileStream(filePath.ToString(), FileMode.Open, FileAccess.Read))
                                {
                                    text = text.Replace($@"#{p}#", filePath.ToString());
                                    run.AddPicture(fileStream, (int)GetPictureType(filePath), $@"{p}", 5300000, 2500000);
                                }
                            }
                            else
                            {
                                if (CommonHelper.IsBase64(filePath.ToString()))
                                {
                                    var fileContent = Convert.FromBase64String(filePath.ToString());
                                    using (var fileStream = new MemoryStream(fileContent))
                                    {
                                        text = text.Replace($@"#{p}#", filePath.ToString());
                                        run.AddPicture(fileStream, (int)GetPictureType(filePath), $@"{p}", 5300000, 2500000);
                                    }
                                }
                                else
                                {
                                    using (var fileStream = await filePath.ToString().GetStreamAsync())
                                    {
                                        text = text.Replace($@"#{p}#", filePath.ToString());
                                        run.AddPicture(fileStream, (int)GetPictureType(filePath), $@"{p}", 5300000, 2500000);
                                    }
                                }

                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                        }
                    }
                }

                if (text.Contains('\n'))
                {
                    run.SetText(string.Empty, 0);

                    var p = text.Split('\n');
                    foreach (var item in p)
                    {
                        run.AppendText(item);
                        run.AddCarriageReturn();
                    }
                }
                else
                {
                    run.SetText(text, 0);

                }
            }
        }


        private static async Task ReplaceKeyObjetAsync(XWPFParagraph para, object model, Func<string, string> func)
        {
            string text = "";
            Type t = model.GetType();
            PropertyInfo[] pi = t.GetProperties();

            foreach (var run in para.Runs)
            {
                text = run.ToString();
                foreach (PropertyInfo p in pi)
                {
                    string key = $"${p.Name}$";
                    var reg = new Regex(@"^#" + p.Name + @"\[\d+,\d+\]#$");
                    var sizeReg = new Regex(@"\[\d+,\d+\]");

                    if (text.Contains(key))
                    {
                        try
                        {
                            text = text.Replace(key, p.GetValue(model, null).ToString());
                        }
                        catch (Exception ex)
                        {
                            text = text.Replace(key, "");
                        }
                    }
                    else
                    {
                        var width = 5300000;
                        var height = 2500000;
                        if (text.Contains($@"#{p.Name}#"))
                        {
                        }
                        else if (reg.IsMatch(text))
                        {
                            var sizeMatch = sizeReg.Match(text);
                            if (sizeMatch.Success)
                            {
                                try
                                {
                                    var w = sizeMatch.Value.Split(',')[0].TrimStart('[');
                                    var h = sizeMatch.Value.Split(',')[1].TrimEnd(']');
                                    width = int.Parse(w) * 9525;
                                    height = int.Parse(h) * 9525;
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("解析图片尺寸错误：" + ex);
                                }

                            }
                        }
                        else
                        {
                            continue;
                        }
                        try
                        {
                            var filePath = p.GetValue(model, null) as string;
                            filePath = func?.Invoke(filePath);
                            if (string.IsNullOrEmpty(filePath))
                            {
                                continue;
                            }
                            if (File.Exists(filePath))
                            {
                                using (var fileStream = new FileStream(filePath.ToString(), FileMode.Open, FileAccess.Read))
                                {
                                    text = "";
                                    run.AddPicture(fileStream, (int)GetPictureType(filePath), $@"{p.Name}", width, height);
                                    NPOI.OpenXmlFormats.Dml.WordProcessing.CT_Inline inline = run.GetCTR().GetDrawingList()[0].inline[0];
                                    var id = (uint)para.Runs.IndexOf(run);
                                    inline.docPr.id = id;
                                }
                            }
                            else
                            {

                                if (CommonHelper.IsBase64(filePath.ToString()))
                                {
                                    var fileContent = Convert.FromBase64String(filePath.ToString());
                                    using (var fileStream = new MemoryStream(fileContent))
                                    {
                                        text = "";
                                        run.AddPicture(fileStream, (int)GetPictureType(filePath), $@"{p}", width, height);
                                        NPOI.OpenXmlFormats.Dml.WordProcessing.CT_Inline inline = run.GetCTR().GetDrawingList()[0].inline[0];
                                        var id = (uint)para.Runs.IndexOf(run);
                                        inline.docPr.id = id;
                                    }
                                }
                                else
                                {
                                    using (var fileStream = await filePath.ToString().GetStreamAsync())
                                    {
                                        text = "";
                                        run.AddPicture(fileStream, (int)GetPictureType(filePath), $@"{p.Name}", width, height);
                                        NPOI.OpenXmlFormats.Dml.WordProcessing.CT_Inline inline = run.GetCTR().GetDrawingList()[0].inline[0];
                                        var id = (uint)para.Runs.IndexOf(run);
                                        inline.docPr.id = id;
                                    }
                                }


                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                        }

                    }
                }
                if (text.Contains('\n'))
                {
                    run.SetText(string.Empty, 0);

                    var p = text.Split('\n');
                    foreach (var item in p)
                    {
                        run.AppendText(item);
                        run.AddCarriageReturn();
                    }
                }
                else
                {
                    run.SetText(text, 0);

                }
            }
        }


        private static PictureType GetPictureType(string str)
        {
            object? result;
            if (!str.StartsWith(Dot))
            {
                var index = str.LastIndexOf(Dot);
                if (index != -1 && str.Length > index + 1)
                {
                    str = str.Substring(index + 1);
                }


            }
            PictureType.TryParse(typeof(PictureType), str.ToUpper(), out result);
            if (result != null)
            {
                return (PictureType)result;
            }
            else
            {
                switch (str.ToUpper())
                {

                    case "JPG":
                        result = PictureType.JPEG;
                        break;

                    default:
                        result = null;
                        break;
                }
                return (PictureType)result;
            }
        }
    }
}