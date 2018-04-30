using System;
using CommandLine;

namespace Openillink.InvoiceGenerator.Models
{
    public class Options
    {
        [Option('i', "input", Default = null, HelpText = "Input file to be processed.", Required = true)]
        public string InputFile { get; set; }

        [Option("from", HelpText = "From date (format yyyy-mm-dd)", Required = false)]
        public DateTime StartDate { get; set; } = DateTime.MinValue;

        [Option("to", HelpText = "To date (format yyyy-mm-dd)", Required = false)]
        public DateTime EndDate { get; set; } = DateTime.Today;
    }
}
