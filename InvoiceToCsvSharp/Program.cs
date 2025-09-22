using System.Threading.Tasks;
using InvoiceToCsvSharp.Utils;

namespace InvoiceToCsvSharp
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            var processor = new InvoiceProcessor();
            await processor.Main();
        }
    }
}
