using System;
using SautinSoft;
using Word.Config;

namespace Word.Converters
{
    public class XlsxConverter : Converter
    {
        private UseOffice SautinSoftOffice { get; }

        public XlsxConverter(UseOffice sautinSoftOffice) : base(".xlsx")
        {
            SautinSoftOffice = sautinSoftOffice;
        }

        protected override void Convert()
        {
            var result = SautinSoftOffice.InitExcel();

            if (result == 0) //succesfully opend program
            {
                do
                {
                    var document = ConversionQueue.Dequeue();

                    string newPath = "";

                    if (document.Name.EndsWith(SupportedExtension))
                    {
                        var newName = document.Name.Replace(SupportedExtension, Constants.PDFExtension);
                        newPath = document.FullPath.Replace(document.Name, newName);
                    }

                    result = SautinSoftOffice.ConvertFile(document.FullPath, newPath, UseOffice.eDirection.XLSX_to_PDF);
                } while (ConversionQueue.Count > 0);

                SautinSoftOffice.CloseExcel();
            }

            ConversionThread.Abort();
        }
    }
}
