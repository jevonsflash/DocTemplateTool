using DocTemplateTool.Common.Helper;
using Flurl.Http;
using PdfSharpCore.Drawing;
using PdfSharpCore.Drawing.Layout;
using PdfSharpCore.Fonts;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.AcroForms;
using PdfSharpCore.Pdf.IO;
using System.Collections;
using System.Reflection;
using System.Text.RegularExpressions;

namespace DocTemplateTool.Pdf
{

    public class Exporter
    {
        const string Dot = ".";

        /// <summary>
        /// 指定对象类型数据源，导出文档
        /// </summary>
        /// <param name="templateStream">模板文档字节流对象</param>
        /// <param name="data">数据源</param>
        /// <param name="func">图片前缀处理委托</param>
        /// <returns>PdfSharp.PdfDocument文档对象</returns>
        public static PdfDocument ExportDocxByObject(Stream templateStream, object data, Func<string, string> func = null)
        {
            var doc = PdfReader.Open(templateStream, PdfDocumentOpenMode.Modify);
 
                ReplaceKeyObjetAsync(doc, data, func);
           

            return doc;
        }

        /// <summary>
        /// 指定对象类型数据源，导出文档
        /// </summary>
        /// <param name="templateFilePath">模板文档路径</param>
        /// <param name="data">数据源</param>
        /// <param name="func">图片前缀处理委托</param>
        /// <returns>PdfSharp.PdfDocument文档对象</returns>
        public static PdfDocument ExportDocxByObject(string templateFilePath, object data, Func<string, string> func = null)
        {
            using (var fileStream = new FileStream(templateFilePath, FileMode.Open))
            {
                return ExportDocxByObject(fileStream, data, func);
            }
        }


        /// <summary>
        /// 指定对象类型数据源，导出文档
        /// </summary>
        /// <param name="templateData">模板文档文件内容</param>
        /// <param name="data">数据源</param>
        /// <param name="func">图片前缀处理委托</param>
        /// <returns>PdfSharp.PdfDocument文档对象</returns>
        public static PdfDocument ExportDocxByObject(byte[] templateData, object data, Func<string, string> func = null)
        {
            using (var stream = new MemoryStream(templateData))
            {
                return ExportDocxByObject(stream, data, func);
            }
        }



        /// <summary>
        /// 指定哈希表类型数据源，导出文档
        /// </summary>
        /// <param name="templateStream">模板文档字节流对象</param>
        /// <param name="data">数据源</param>
        /// <param name="func">图片前缀处理委托</param>
        /// <returns>PdfSharp.PdfDocument文档对象</returns>
        public static PdfDocument ExportDocxByDictionary(Stream templateStream, IDictionary<string, object> data, Func<string, string> func = null)
        {
            var doc = PdfReader.Open(templateStream, PdfDocumentOpenMode.Modify);

            ReplaceKeyDictionaryAsync(doc, data, func);


            return doc;
        }

        /// <summary>
        /// 指定哈希表类型数据源，导出文档
        /// </summary>
        /// <param name="templateFilePath">模板文档路径</param>
        /// <param name="data">数据源</param>
        /// <param name="func">图片前缀处理委托</param>
        /// <returns>PdfSharp.PdfDocument文档对象</returns>

        public static PdfDocument ExportDocxByDictionary(string templateFilePath, IDictionary<string, object> data, Func<string, string> func = null)
        {
            using (var fileStream = new FileStream(templateFilePath, FileMode.Open))
            {
                return ExportDocxByDictionary(fileStream, data, func);
            }
        }


        /// <summary>
        /// 指定哈希表类型数据源，导出文档
        /// </summary>
        /// <param name="templateData">模板文档文件内容</param>
        /// <param name="data">数据源</param>
        /// <param name="func">图片前缀处理委托</param>
        /// <returns>PdfSharp.PdfDocument文档对象</returns>

        public static PdfDocument ExportDocxByDictionary(byte[] templateData, IDictionary<string, object> data, Func<string, string> func = null)
        {
            using (var stream = new MemoryStream(templateData))
            {
                return ExportDocxByDictionary(stream, data, func);
            }
        }


        private static async Task ReplaceKeyDictionaryAsync(PdfDocument doc, IDictionary<string, object> data, Func<string, string> func)
        {
            PdfAcroForm form = doc.AcroForm;
            if (form.Elements.ContainsKey("/NeedAppearances"))
            {
                form.Elements["/NeedAppearances"] = new PdfBoolean(true);
            }
            else
            {
                form.Elements.Add("/NeedAppearances", new PdfBoolean(true));
            }
            string text = "";
           
            var paramsReg = new Regex(@"\[\d+\]");

            foreach (var fieldName in form.Fields.Names)
            {
                var run = form.Fields[fieldName] as PdfTextField;
                text = run.Name;
                XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode);

                XFont font = new XFont(GlobalFontSettings.FontResolver.DefaultFontName, 18, XFontStyle.Regular, options);
                foreach (var p in data.Keys)
                {
                    //string key = $"${p.Name}$";
                    var textReg = new Regex(@"\$" + p + @"(\[\w+\])?\$");
                    var assetReg = new Regex(@"#" + p + @"(\[\d+,\d+\])?#");

                    if (textReg.IsMatch(text))
                    {
                        try
                        {
                            var value = data[p];
                            if (value is IEnumerable && value is not string)
                            {

                            }
                            else
                            {
                                var stringValue = value.ToString();
                                if (stringValue.Contains('\n'))
                                {
                                    run.MultiLine = true;

                                }

                                run.Value = new PdfString(stringValue, PdfStringEncoding.Unicode);
                                run.ReadOnly = true;
                            }
                        }
                        catch (Exception ex)
                        {
                            text = "";
                        }
                    }
                    else if (assetReg.IsMatch(text))
                    {
                        try
                        {
                            var filePath = data[p] as string;
                            filePath = func?.Invoke(filePath);
                            if (string.IsNullOrEmpty(filePath))
                            {
                                continue;
                            }
                            if (File.Exists(filePath))
                            {
                                using (var fileStream = new FileStream(filePath.ToString(), FileMode.Open, FileAccess.Read))
                                {
                                    var rectangle = run.Elements.GetRectangle("/Rect");
                                    var xForm = new XForm(doc, rectangle.Size);
                                    using (var xGraphics = XGraphics.FromPdfPage(doc.Pages[0]))
                                    {
                                        var image = XImage.FromStream(() => fileStream);
                                        xGraphics.DrawImage(image, rectangle.ToXRect() + new XPoint(0, 400));
                                        var state = xGraphics.Save();
                                        xGraphics.Restore(state);
                                    }
                                }
                            }
                            else
                            {

                                if (!CommonHelper.IsUrl(filePath))
                                {
                                    string pattern = @"data:image/(?<type>.+?);base64,(?<data>[^""]+)";
                                    Regex regex = new Regex(pattern, RegexOptions.Compiled);
                                    var match = regex.Match(filePath);
                                    int index = 0;

                                    var type = match.Groups["type"].Value;
                                    var dataString = match.Groups["data"].Value;


                                    var fileContent = Convert.FromBase64String(dataString);
                                    using (var fileStream = new MemoryStream(fileContent))
                                    {
                                        var rectangle = run.Elements.GetRectangle("/Rect");
                                        var xForm = new XForm(doc, rectangle.Size);
                                        using (var xGraphics = XGraphics.FromPdfPage(doc.Pages[0]))
                                        {
                                            var image = XImage.FromStream(() => fileStream);
                                            xGraphics.DrawImage(image, rectangle.ToXRect() + new XPoint(0, 400));
                                            var state = xGraphics.Save();
                                            xGraphics.Restore(state);
                                        }
                                    }
                                }
                                else
                                {
                                    using (var fileStream = filePath.ToString().GetStreamAsync().Result)
                                    {
                                        var rectangle = run.Elements.GetRectangle("/Rect");
                                        var xForm = new XForm(doc, rectangle.Size);
                                        using (var xGraphics = XGraphics.FromPdfPage(doc.Pages[0]))
                                        {
                                            var image = XImage.FromStream(() => fileStream);
                                            xGraphics.DrawImage(image, rectangle.ToXRect() + new XPoint(0, 400));
                                            var state = xGraphics.Save();
                                            xGraphics.Restore(state);
                                        }
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

            }
        }


        private static async Task ReplaceKeyObjetAsync(PdfDocument doc, object model, Func<string, string> func)
        {
            PdfAcroForm form = doc.AcroForm;
            if (form.Elements.ContainsKey("/NeedAppearances"))
            {
                form.Elements["/NeedAppearances"] = new PdfBoolean(true);
            }
            else
            {
                form.Elements.Add("/NeedAppearances", new PdfBoolean(true));
            }
            string text = "";
            Type t = model.GetType();
            PropertyInfo[] pi = t.GetProperties();
            var paramsReg = new Regex(@"\[\d+\]");

            foreach (var fieldName in form.Fields.Names)
            {
                var run = form.Fields[fieldName] as PdfTextField;
                text = run.Name;
                XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode);

                XFont font = new XFont(GlobalFontSettings.FontResolver.DefaultFontName, 18, XFontStyle.Regular, options);
                foreach (PropertyInfo p in pi)
                {
                    //string key = $"${p.Name}$";
                    var textReg = new Regex(@"\$" + p.Name + @"(\[\w+\])?\$");
                    var assetReg = new Regex(@"#" + p.Name + @"(\[\d+,\d+\])?#");

                    if (textReg.IsMatch(text))
                    {
                        try
                        {
                            var value = p.GetValue(model, null);
                            if (value is IEnumerable && value is not string)
                            {
                               
                            }
                            else
                            {
                                var stringValue = value.ToString();
                                if (stringValue.Contains('\n'))
                                {
                                    run.MultiLine = true;

                                }

                                run.Value = new PdfString(stringValue, PdfStringEncoding.Unicode);
                                run.ReadOnly = true;
                            }
                        }
                        catch (Exception ex)
                        {
                            text = "";
                        }
                    }
                    else if (assetReg.IsMatch(text))
                    {                
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
                                    var rectangle = run.Elements.GetRectangle("/Rect");
                                    var xForm = new XForm(doc, rectangle.Size);
                                    using (var xGraphics = XGraphics.FromPdfPage(doc.Pages[0]))
                                    {
                                        var image = XImage.FromStream(() => fileStream);
                                        xGraphics.DrawImage(image, rectangle.ToXRect() + new XPoint(0, 400));
                                        var state = xGraphics.Save();
                                        xGraphics.Restore(state);
                                    }
                                }
                            }
                            else
                            {

                                if (!CommonHelper.IsUrl(filePath))
                                {
                                    string pattern = @"data:image/(?<type>.+?);base64,(?<data>[^""]+)";
                                    Regex regex = new Regex(pattern, RegexOptions.Compiled);
                                    var match = regex.Match(filePath);
                                    int index = 0;

                                    var type = match.Groups["type"].Value;
                                    var dataString = match.Groups["data"].Value;


                                    var fileContent = Convert.FromBase64String(dataString);
                                    using (var fileStream = new MemoryStream(fileContent))
                                    {
                                        var rectangle = run.Elements.GetRectangle("/Rect");
                                        var xForm = new XForm(doc, rectangle.Size);
                                        using (var xGraphics = XGraphics.FromPdfPage(doc.Pages[0]))
                                        {
                                            var image = XImage.FromStream(() => fileStream);
                                            xGraphics.DrawImage(image, rectangle.ToXRect() + new XPoint(0, 400));
                                            var state = xGraphics.Save();
                                            xGraphics.Restore(state);
                                        }
                                    }
                                }
                                else
                                {
                                    using (var fileStream = filePath.ToString().GetStreamAsync().Result)
                                    {
                                        var rectangle = run.Elements.GetRectangle("/Rect");
                                        var xForm = new XForm(doc, rectangle.Size);
                                        using (var xGraphics = XGraphics.FromPdfPage(doc.Pages[0]))
                                        {
                                            var image = XImage.FromStream(() => fileStream);
                                            xGraphics.DrawImage(image, rectangle.ToXRect() + new XPoint(0, 400));
                                            var state = xGraphics.Save();
                                            xGraphics.Restore(state);
                                        }
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
                
            }
        }


        private static string? GetStringValue(object value)
        {
            if (value == null) return "";
            if (value is DateTime)
            {
                return ((DateTime)value).ToString("yyyy-MM-dd hh:mm");

            }
            else if (value is double)
            {
                return ((double)value).ToString("0.00");
            }
            else if (value is float)
            {
                return ((float)value).ToString("0.00");
            }

            return value.ToString();

        }

    }

}