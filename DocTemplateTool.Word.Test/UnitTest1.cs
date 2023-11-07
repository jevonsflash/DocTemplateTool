using System.Reflection;
using System.Text;
using DocTemplateTool.Word.Test.Model;
using NPOI.SS.Formula.Functions;
using static System.Net.Mime.MediaTypeNames;

namespace DocTemplateTool.Word.Test
{
    [TestClass]
    public class UnitTest1
    {


        private static readonly string templatePath_Doc = Path.Combine(new FileInfo(Assembly.GetExecutingAssembly().Location).Directory.FullName, "Assets");

        private static readonly string outputPath_Doc = new FileInfo(Assembly.GetExecutingAssembly().Location).Directory.FullName;

        //可以自定义目标文件的路径
        //private static readonly string outputPath_Doc = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);



        /// <summary>
        /// 测试：心电报告
        /// </summary>
        [TestMethod]
        public void TestMethod1()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.WriteLine("Generator begin");
            byte[] docFileContent;

            var docinfo = GetReportDocInfo();

            var result = Word.Exporter.ExportDocxByObject(Path.Combine(templatePath_Doc, $"ReportTemplate.docx"), docinfo, (s) => s);

            using (var memoryStream = new MemoryStream())
            {
                result.Write(memoryStream);
                memoryStream.Seek(0, SeekOrigin.Begin);
                docFileContent = memoryStream.ToArray();
            }

            File.WriteAllBytes(Path.Combine(outputPath_Doc, $"Report.docx"), docFileContent);


        }

        /// <summary>
        /// 测试：处方笺
        /// </summary>
        [TestMethod]
        public void TestMethod2()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.WriteLine("Generator begin");
            byte[] docFileContent;

            var docinfo = GetReportDocInfo();

            var result = Word.Exporter.ExportDocxByObject(Path.Combine(templatePath_Doc, $"RecipeTemplate.docx"), docinfo, (s) => s);

            using (var memoryStream = new MemoryStream())
            {
                result.Write(memoryStream);
                memoryStream.Seek(0, SeekOrigin.Begin);
                docFileContent = memoryStream.ToArray();
            }

            File.WriteAllBytes(Path.Combine(outputPath_Doc, $"Recipe.docx"), docFileContent);


        }

        /// <summary>
        /// 测试：员工健康信息报告
        /// </summary>
        [TestMethod]
        public void TestMethod3()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.WriteLine("Generator begin");
            byte[] docFileContent;

            var docinfo = GetHealthReportDocInfo();

            var result = Word.Exporter.ExportDocxByObject(Path.Combine(templatePath_Doc, $"ReportTemplate2.docx"), docinfo, (s) => s);

            using (var memoryStream = new MemoryStream())
            {
                result.Write(memoryStream);
                memoryStream.Seek(0, SeekOrigin.Begin);
                docFileContent = memoryStream.ToArray();
            }


            File.WriteAllBytes(Path.Combine(outputPath_Doc, $"Report2.docx"), docFileContent);


        }


        /// <summary>
        /// 测试：工资登记表
        /// </summary>

        [TestMethod]
        public void TestMethod4()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.WriteLine("Generator begin");
            byte[] docFileContent;

            var docinfo = new
            {

                Dept = "研发部",
                Date = DateTime.Now,
                Details = new List<dynamic>() {

                  new
                  {
                      Number = "mk201406gz324",
                    Name = "张三",
                    BaseAccount = 2000,
                    TechnicalAllowance = 1500,
                    SeniorityAllowance = 700,
                    DutyAllowance = 300,
                    JobAllowance = 200,
                    ItemSum = 4700
                }
,
new
{
                Number = "mk201406gz325",
                      Name = "李四",
                      BaseAccount = 2500,
                      TechnicalAllowance = 2000,
                      SeniorityAllowance = 800,
                      DutyAllowance = 350,
                      JobAllowance = 250,
                      ItemSum = 5800
                  }
            ,
            new
{
                Number = "mk201406gz326",
                      Name = "王五",
                      BaseAccount = 1800,
                      TechnicalAllowance = 1200,
                      SeniorityAllowance = 600,
                      DutyAllowance = 300,
                      JobAllowance = 200,
                      ItemSum = 4900
                  }
            ,
            new
{
                Number = "mk201406gz327",
                      Name = "赵六",
                      BaseAccount = 2200,
                      TechnicalAllowance = 1600,
                      SeniorityAllowance = 750,
                      DutyAllowance = 350,
                      JobAllowance = 250,
                      ItemSum = 5650
                  }
            ,
            new
{
                Number = "mk201406gz328",
                      Name = "孙策",
                      BaseAccount = 1950,
                      TechnicalAllowance = 1350,
                      SeniorityAllowance = 650,
                      DutyAllowance = 325,
                      JobAllowance = 225,
                      ItemSum = 5550
                  }
            ,
            new
{
                Number = "mk201406gz329",
                      Name = "周瑜",
                      BaseAccount = 2350,
                      TechnicalAllowance = 1750,
                      SeniorityAllowance = 850,
                      DutyAllowance = 375,
                      JobAllowance = 275,
                      ItemSum = 6450
                  }

        },
                DeptorSum = 50000,
                LenderSum = 50000,
                //12800   9400    4350    2000    1400    33050

                Sum1 = 12800,
                Sum2 = 9400,
                Sum3 = 4350,
                Sum4 = 2000,
                Sum5 = 1400,
                Sum = 33050,
                Auditor = "王五",
                Register = "赵六",

            };

            var result = Word.Exporter.ExportDocxByObject(Path.Combine(templatePath_Doc, $"SalaryTemplate.docx"), docinfo, (s) => s);

            using (var memoryStream = new MemoryStream())
            {
                result.Write(memoryStream);
                memoryStream.Seek(0, SeekOrigin.Begin);
                docFileContent = memoryStream.ToArray();
            }


            File.WriteAllBytes(Path.Combine(outputPath_Doc, $"Salary.docx"), docFileContent);


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