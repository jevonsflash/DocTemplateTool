using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using SautinSoft;
using Word.Config;
using Word.Converters;
using Word.Models;

namespace Word
{
    public class Service
    {
        private static FileSystemWatcher _watcher;
        private ConverterQueue _converterQueue;
        private IList<Converter> _converters;

        public void Start()
        {
            BootstrapApplication();

            var folderPath = ConfigurationManager.AppSettings["WatchFolder"];
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                throw new KeyNotFoundException("No WatchFolder specified in appSettings");
            }

            _watcher = new FileSystemWatcher();
            _watcher.Path = folderPath;
            _watcher.Created += OnChanged;
            _watcher.Filter = "*.*";
            _watcher.Changed += new FileSystemEventHandler(OnChanged);
            _watcher.EnableRaisingEvents = true;

        }

        public void Stop()
        {
            _watcher?.Dispose();
        }

        public void BootstrapApplication()
        {
            _converters = new List<Converter>
            {
                new DocxConverter(new UseOffice()),
                new PptxConverter(new UseOffice()),
                new XlsxConverter(new UseOffice())
            };

            _converterQueue = new ConverterQueue(_converters);

        }

        private void OnChanged(object sender, FileSystemEventArgs e)
        {
            if (!ValidateFiletype(e.Name)) return;

            var doc = new DocumentInfo()
            {
                Name = e.Name,
                Path = e.FullPath.TrimEnd(e.Name.ToCharArray()),
                FullPath = e.FullPath,
                Extension = "." + e.FullPath.Split('.').Last()
            };

            _converterQueue.Push(doc);
        }

        private bool ValidateFiletype(string fileName)
        {
            var extension = fileName.Split('.').Last();

            if (fileName.Contains('~'))
            {
                return false;
            }

            if (AppConfig.AllowedFileTypes.Contains(extension))
            {
                return true;
            }

            return false;
        }
    }
}
