using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Word.Models;

namespace Word.Converters
{
    public abstract class Converter
    {
        // Fields
        public readonly string SupportedExtension;

        // Properties
        protected Queue<DocumentInfo> ConversionQueue { get; }
        protected Thread ConversionThread { get; set; }

        // Constructors
        protected Converter()
        {
            ConversionQueue = new Queue<DocumentInfo>();
        }

        protected Converter(string extension) : this()
        {
            SupportedExtension = extension;
        }

        // Methods
        public void Push(DocumentInfo document)
        {
            if (ConversionQueue.All(x => x.Name != document.Name))
            {
                ConversionQueue.Enqueue(document);
            }

            if (ConversionThread == null || !ConversionThread.IsAlive)
            {
                ConversionThread = new Thread(Convert);
                ConversionThread.Start();
            }
        }

        protected abstract void Convert();

    }

}