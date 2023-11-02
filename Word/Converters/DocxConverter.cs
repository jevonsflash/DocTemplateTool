using SautinSoft;
using Word.Config;

namespace Word.Converters
{
    public class DocxConverter : Converter
    {
        private UseOffice SautinSoftOffice { get; }

        public DocxConverter(UseOffice sautinSoftOffice) : base(".docx")
        {
            SautinSoftOffice = sautinSoftOffice;
        }

        protected override void Convert()
        {
            var result = SautinSoftOffice.InitWord();

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

                    result = SautinSoftOffice.ConvertFile(document.FullPath, newPath, UseOffice.eDirection.DOCX_to_PDF);
                } while (ConversionQueue.Count > 0);

                SautinSoftOffice.CloseWord();
            }

            ConversionThread.Abort();
        }
    }
}
