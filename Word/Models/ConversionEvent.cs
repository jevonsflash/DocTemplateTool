namespace Word.Models
{
    public class ConversionEvent
    {
        public DocumentInfo DocumentInfo { get; set; }
        public string OutputExtension { get; set; }
        public string OutputFullPath { get; set; }
    }

    public class DocumentInfo
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public string FullPath { get; set; }
        public string Extension { get; set; }
    }
}
