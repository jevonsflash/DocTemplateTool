using System.Collections.Generic;
using DocTemplateTool.Word.Converters;
using DocTemplateTool.Word.Models;

namespace DocTemplateTool.Word
{
    public class ConverterQueue
    {
        public IDictionary<string, Converter> Converters { get; }

        // constructor injection dependencies
        public ConverterQueue(IEnumerable<Converter> converters)
        {
            Converters = new Dictionary<string, Converter>();

            foreach (var converter in converters)
            {
                Converters.Add(converter.SupportedExtension, converter);
            }
        }

        public void Push(DocumentInfo documentInfo)
        {
            Converters[documentInfo.Extension].Push(documentInfo);
        }
    }
}
