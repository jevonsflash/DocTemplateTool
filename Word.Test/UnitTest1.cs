using System.Reflection;
using System.Text;
using NPOI.SS.Formula.Functions;
using Word.Test.Model;
using static System.Net.Mime.MediaTypeNames;

namespace Word.Test
{
    [TestClass]
    public class UnitTest1
    {


        private static readonly string templatePath_Doc = System.IO.Path.Combine(new FileInfo(Assembly.GetExecutingAssembly().Location).Directory.FullName, "Assets");

        private static readonly string outputPath_Doc = new FileInfo(Assembly.GetExecutingAssembly().Location).Directory.FullName;

        [TestMethod]
        public void TestMethod1()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.WriteLine("Generator begin");
            byte[] docFileContent;

            var docinfo = GetReportDocInfo();
            var ext = ".docx";
            var result = Word.Exporter.ExportDocxByObject(Path.Combine(templatePath_Doc, $"ReportTemplate.docx"), docinfo, (s) => s);

            using (var memoryStream = new MemoryStream())
            {
                result.Write(memoryStream);
                memoryStream.Seek(0, SeekOrigin.Begin);
                docFileContent = memoryStream.ToArray();
            }

            File.WriteAllBytes(Path.Combine(outputPath_Doc, $"Report.docx"), docFileContent);


        }


        [TestMethod]
        public void TestMethod2()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.WriteLine("Generator begin");
            byte[] docFileContent;

            var docinfo = GetReportDocInfo();
            var ext = ".docx";
            var result = Word.Exporter.ExportDocxByObject(Path.Combine(templatePath_Doc, $"RecipeTemplate.docx"), docinfo, (s) => s);

            using (var memoryStream = new MemoryStream())
            {
                result.Write(memoryStream);
                memoryStream.Seek(0, SeekOrigin.Begin);
                docFileContent = memoryStream.ToArray();
            }

            File.WriteAllBytes(Path.Combine(outputPath_Doc, $"Recipe.docx"), docFileContent);


        }

        [TestMethod]
        public void TestMethod3()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.WriteLine("Generator begin");
            byte[] docFileContent;

            var docinfo = GetHealthReportDocInfo();
            var ext = ".docx";
            var result = Word.Exporter.ExportDocxByObject(Path.Combine(templatePath_Doc, $"ReportTemplate2.docx"), docinfo, (s) => s);

            using (var memoryStream = new MemoryStream())
            {
                result.Write(memoryStream);
                memoryStream.Seek(0, SeekOrigin.Begin);
                docFileContent = memoryStream.ToArray();
            }


            File.WriteAllBytes(Path.Combine(outputPath_Doc, $"Report2.docx"), docFileContent);


        }




        protected virtual ReportDocInfo GetReportDocInfo()
        {
            var docinfo = new ReportDocInfo();

            docinfo.Id = 202247589;
            docinfo.HospitalName = "广东省某互联网医院";
            docinfo.ReportTime = new DateTime(2023, 11, 1);
            docinfo.DepartmentName = "心血管内科";
            docinfo.AuditEmployeeName = "王五";
            docinfo.AuditEmployeeSignature = System.IO.Path.Combine(new FileInfo(Assembly.GetExecutingAssembly().Location).Directory.FullName, "Assets", $"TestPic.jpg");
            docinfo.DraftEmployeeName = "李四";
            docinfo.DraftEmployeeSignature = System.IO.Path.Combine(new FileInfo(Assembly.GetExecutingAssembly().Location).Directory.FullName, "Assets", $"TestPic.jpg");
            docinfo.ClientName = "张三";
            docinfo.ClientAge = "35";
            docinfo.Name = "良性高血压的处方";
            docinfo.Diagnostic = "良性高血压";
            docinfo.ClientSex = "男";
            docinfo.Price = 12.0M;
            docinfo.Graphic = System.IO.Path.Combine(new FileInfo(Assembly.GetExecutingAssembly().Location).Directory.FullName, "Assets", $"ECGGraphic.png");
            docinfo.RpList = new List<string>()
            {
                "1.苯磺酸氨氯地平片(京新)(省采)\n 规格：2mg*28片\t×10片\n用法用量：2片/次，每日三次，口服。",
                "2.苯磺酸氨氯地平片(络活喜)(非省采)\n 规格：5mg*28片\t×20片\n用法用量：1片/次，每日两次，口服。",
            };
            docinfo.Rps = string.Join('\n', docinfo.RpList);
            return docinfo;
        }


        protected virtual HealthReportDocInfo GetHealthReportDocInfo()
        {
            var docinfo = new HealthReportDocInfo();

            docinfo.ClientName = "XX科技股份有限公司";
            docinfo.Title = "健康企业员工健康管理周报";
            docinfo.TimeSpan = "2023年10月19日-10月29日";
            docinfo.Count1 = 2; docinfo.Count2 = 3;
            docinfo.TotalMemberCount = 60;
            docinfo.BloodPressureTestMemberCount = 42;
            docinfo.BloodGlucoseTestMemberCount = 43;
            docinfo.ECGTestMemberCount = 30;
            docinfo.BloodPressureAnalysis = new List<AnalysisList>()
            {
                 new AnalysisList()
                 {

                      Dept="技术部", Count=8
                 },
                  new AnalysisList()
                 {

                      Dept="总经办", Count=1
                 },
                   new AnalysisList()
                 {

                      Dept="客户部", Count=2
                 }


            };

            docinfo.BloodGlucoseAnalysis = new List<AnalysisList>()
            {
                 new AnalysisList()
                 {

                      Dept="技术部", Count=4
                 },
                  new AnalysisList()
                 {

                      Dept="总经办", Count=1
                 },
                   new AnalysisList()
                 {

                      Dept="客户部", Count=1
                 }


            };

            docinfo.ECGAnalysis = new List<AnalysisList>()
            {
                 new AnalysisList()
                 {

                      Dept="技术部", Count=4
                 },
                  new AnalysisList()
                 {

                      Dept="总经办", Count=1
                 },
                   new AnalysisList()
                 {

                      Dept="客户部", Count=1
                 }


            };

            docinfo.BloodPressureList = new List<DetailList> { new DetailList() {
            Dept="技术部",
             Name="张三",
             Value="144/66",
             Result="↑"
            },
            new DetailList() {
            Dept="技术部",
             Name="李四",
             Value="150/86",
             Result="↑"
            },
             new DetailList() {
            Dept="技术部",
             Name="张伟",
             Value="149/86",
             Result="↑"
            },
            new DetailList() {
            Dept="技术部",
             Name="李青",
             Value="128/92",
             Result="↑"
            }
            };
            docinfo.BloodGlucoseList = new List<DetailList> { new DetailList() {
            Dept="技术部",
             Name="张伟",
             Value="6.3",
             Result="↑"
            },
            new DetailList() {
            Dept="技术部",
             Name="王芳",
             Value="6.5",
             Result="↑"
            }
            };
            docinfo.ECGList = new List<DetailList> { new DetailList() {
            Dept="技术部",
             Name="张敏",
             Value="122",
             Result="↑"
            },
            new DetailList() {
            Dept="技术部",
             Name="陈婷",
             Value="55",
             Result="↓"
            }
            };
            return docinfo;
        }

    }
}