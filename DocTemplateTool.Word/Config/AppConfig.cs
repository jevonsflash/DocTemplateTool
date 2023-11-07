using System.Configuration;

namespace DocTemplateTool.Word.Config
{
    public class AppConfig
    {
        public static string[] AllowedFileTypes = ConfigurationManager.AppSettings["SupportedFileTypes"].Split(';');
    }
}
