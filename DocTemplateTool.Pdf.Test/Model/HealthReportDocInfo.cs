using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTemplateTool.Pdf.Test.Model
{
    public class HealthReportDocInfo
    {

        public string ClientName { get; set; }
        public string Title { get; set; }
        public string TimeSpan { get; set; }
        public int Count1 { get; set; }
        public int Count2 { get; set; }
        public int TotalMemberCount { get; set; }
        public int BloodGlucoseTestMemberCount { get; set; }
        public int BloodPressureTestMemberCount { get; set; }
        public int ECGTestMemberCount { get; set; }

        public int BloodGlucoseTestPassCount { get; set; }
        public int BloodPressureTestPassCount { get; set; }
        public int ECGTestPassCount { get; set; }

        public int BloodGlucoseTestFailedCount => BloodGlucoseTestMemberCount - BloodGlucoseTestPassCount;
        public int BloodPressureTestFailedCount => BloodPressureTestMemberCount - BloodPressureTestPassCount;
        public int ECGTestFailedCount => ECGTestMemberCount - ECGTestPassCount;

        public double BloodGlucoseTestPassRatio => BloodGlucoseTestPassCount / BloodGlucoseTestMemberCount;
        public double BloodPressureTestPassRatio => BloodPressureTestPassCount / BloodPressureTestMemberCount;
        public double ECGTestPassRatio => ECGTestPassCount / ECGTestMemberCount;
        public List<AnalysisList> BloodPressureAnalysis { get; set; }
        public List<AnalysisList> BloodGlucoseAnalysis { get; set; }
        public List<AnalysisList> ECGAnalysis { get; set; }
        public List<DetailList> BloodPressureList { get; set; }
        public List<DetailList> BloodGlucoseList { get; set; }
        public List<DetailList> ECGList { get; set; }

    }
}
