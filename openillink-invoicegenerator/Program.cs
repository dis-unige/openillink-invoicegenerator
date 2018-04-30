using CommandLine;
using Openillink.InvoiceGenerator.Models;

namespace Openillink.InvoiceGenerator
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed(options =>
                {
                    var generator = new InvoiceGenerator();
                    generator.Run(options);
                });
        }
    }
}
