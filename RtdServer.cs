using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration.Rtd;

namespace ExcelDnaRTDSample
{
    [ComVisible(true)]
    [ProgId("ExcelDnaRtdSample.RtdServer")]
    [Guid("BC089024-902F-4E44-9F2D-C62325107655")]
    public class RtdServer : IRtdServer
    {
        public int ServerStart(IRTDUpdateEvent CallbackObject)
        {
            return 1;
        }

        public object ConnectData(int topicId, ref Array strings, ref bool newValues)
        {
            return 1;
        }

        public Array RefreshData(ref int topicCount)
        {
            return new[] {1};
        }

        public void DisconnectData(int topicID)
        {

        }

        public int Heartbeat()
        {
            return 1;
        }

        public void ServerTerminate()
        {
        }

        public string GetValue()
        {
            return nameof(RtdServer);
        }
    }
}