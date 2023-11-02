using System.Collections.Generic;
using Word.Converters;
using Word.Models;

namespace Word
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
