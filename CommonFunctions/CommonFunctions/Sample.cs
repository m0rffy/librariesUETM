using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonFunctions
{
    public class Sample
    {
        SampleObject[] Samples;
    }

    public class SampleObject
    {
        public SampleObject(int SampleCount, long Amplitude, int SynchStatus, string DeviceName)
        {
            this.SampleCount = SampleCount;
            this.Amplitude = Amplitude;
            this.SynchStatus = SynchStatus;
            this.DeviceName = DeviceName;
        }
        public int SampleCount = 0;
        public long Amplitude = 0;
        public int SynchStatus = 0;
        public string DeviceName = "";
    }
}
