using DocTemplateTool.Word.Config;
using SautinSoft;

namespace DocTemplateTool.Word.Converters
{
    public class PptxConverter : Converter
    {
        private UseOffice SautinSoftOffice { get; }

        public PptxConverter(UseOffice sautinSoftOffice) : base(".pptx")
        {
            SautinSoftOffice = sautinSoftOffice;
        }

        protected override void Convert()
        {
            var result = SautinSoftOffice.InitPowerPoint();

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

                    result = SautinSoftOffice.ConvertFile(document.FullPath, newPath, UseOffice.eDirection.PPTX_to_PDF);
                } while (ConversionQueue.Count > 0);

                SautinSoftOffice.ClosePowerPoint();
            }

            ConversionThread.Abort();
        }
    }
}
