using System.Reflection;
using System.Text;
using DocTemplateTool.Pdf.Test.Model;
using DocTemplateTool.Pdf;
using static System.Net.Mime.MediaTypeNames;

namespace DocTemplateTool.Pdf.Test
{
    [TestClass]
    public class UnitTest1
    {


        private static readonly string templatePath_Doc = Path.Combine(new FileInfo(Assembly.GetExecutingAssembly().Location).Directory.FullName, "Assets");

        private static readonly string outputPath_Doc = new FileInfo(Assembly.GetExecutingAssembly().Location).Directory.FullName;

        //可以自定义目标文件的路径
        //private static readonly string outputPath_Doc = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);



    
        /// <summary>
        /// 测试：处方笺
        /// </summary>
        [TestMethod]
        public void TestMethod1()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.WriteLine("Generator begin");
            var docinfo = GetReportDocInfo();

            var result = Exporter.ExportDocxByObject(Path.Combine(templatePath_Doc, $"RecipeTemplate.pdf"), docinfo, (s) => s);
            result.Save(Path.Combine(outputPath_Doc, $"Recipe.pdf"));

        }



        protected virtual ReportDocInfo GetReportDocInfo()
        {
            var docinfo = new ReportDocInfo
            {
                Id = 202247589,
                HospitalName = "广东省某互联网医院",
                ReportTime = new DateTime(2023, 11, 1),
                DepartmentName = "心血管内科",
                AuditEmployeeName = "王五",
                AuditEmployeeSignature = Path.Combine(new FileInfo(Assembly.GetExecutingAssembly().Location).Directory.FullName, "Assets", $"TestPic.jpg"),
                DraftEmployeeName = "李四",
                DraftEmployeeSignature = Path.Combine(new FileInfo(Assembly.GetExecutingAssembly().Location).Directory.FullName, "Assets", $"TestPic.jpg"),
                ClientName = "张三",
                ClientAge = "35",
                Name = "良性高血压的处方",
                Diagnostic = "良性高血压",
                ClientSex = "男",
                Price = 12.0M,
                Graphic = Path.Combine(new FileInfo(Assembly.GetExecutingAssembly().Location).Directory.FullName, "Assets", $"ECGGraphic.png"),
                RpList = new List<string>()
                {
                    "1.苯磺酸氨氯地平片(京新)(省采)\n 规格：2mg*28片\t×10片\n用法用量：2片/次，每日三次，口服。",
                    "2.苯磺酸氨氯地平片(络活喜)(非省采)\n 规格：5mg*28片\t×20片\n用法用量：1片/次，每日两次，口服。",
                }
            };
            docinfo.Rps = string.Join('\n', docinfo.RpList);
            return docinfo;
        }



    }
}