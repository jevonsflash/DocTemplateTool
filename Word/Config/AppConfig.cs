using System.Configuration;

namespace Word.Config
{
    public class AppConfig
    {
        public static string[] AllowedFileTypes = ConfigurationManager.AppSettings["SupportedFileTypes"].Split(';');
    }
}
