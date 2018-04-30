using FileHelpers;

namespace Openillink.InvoiceGenerator.Helpers
{
    public class StringConverter : ConverterBase
    {
        public override object StringToField(string from)
        {
            return from.FixEncoding();
        }
    }
}
