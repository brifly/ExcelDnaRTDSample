using System.Runtime.InteropServices;

namespace ExcelDnaRTDSample
{
    [ComVisible(true)]
    public class NonRtdServer
    {
        public string GetValue()
        {
            return nameof(NonRtdServer);
        }
    }
}