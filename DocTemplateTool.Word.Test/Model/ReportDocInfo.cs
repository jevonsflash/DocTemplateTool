using System;

namespace DocTemplateTool.Word.Test.Model
{
    public class ReportDocInfo
    {
        public long Id { get; set; }

        public long OrganizationUnitExtensionId { get; set; }

        public string HospitalName { get; set; }
        public string DepartmentName { get; set; }
        public string ClientName { get; set; }
        public string ClientAge { get; set; }
        public string ClientSex { get; set; }
        public DateTime ReportTime { get; set; }
        public string ReportTimeString => ReportTime.ToString("yyyy-MM-dd HH:mm");

        public string AuditEmployeeName { get; set; }

        public string AuditEmployeeNumber { get; set; }

        public string DraftEmployeeName { get; set; }

        public string DraftEmployeeNumber { get; set; }

        public string DraftEmployeeSignature { get; set; }

        public string AuditEmployeeSignature { get; set; }

        public string Name { get; set; }
        public string Graphic { get; set; }
        public string Diagnostic { get; set; }


        public string Rps { get; set; }
        public List<string> RpList { get; set; }

        public decimal Price { get; set; }
    }
}