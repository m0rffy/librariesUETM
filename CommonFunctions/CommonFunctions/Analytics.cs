using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonFunctions
{
    public class Analytics
    {
        AnalyticsObject[] AnalyticsObjects;
    }

    public class AnalyticsObject
    {
        public AnalyticsObject(double EtalonCurrent, double Channel1AmplitudeError, double Channel1PhaseError, double Channel2AmplitudeError, double Channel2PhaseError, double Channel3AmplitudeError, double Channel3PhaseError, double Channel4AmplitudeError, double Channel4PhaseError)
        {
            this.EtalonCurrent = EtalonCurrent;
            this.Channel1AmplitudeError = Channel1AmplitudeError;
            this.Channel1PhaseError = Channel1PhaseError;
            this.Channel2AmplitudeError = Channel2AmplitudeError;
            this.Channel2PhaseError = Channel2PhaseError;
            this.Channel3AmplitudeError = Channel3AmplitudeError;
            this.Channel3PhaseError = Channel3PhaseError;
            this.Channel4AmplitudeError = Channel4AmplitudeError;
            this.Channel4PhaseError = Channel4PhaseError;
        }

        public double EtalonCurrent = 0;
        public double Channel1AmplitudeError = 0;
        public double Channel1PhaseError = 0;
        public double Channel2AmplitudeError = 0;
        public double Channel2PhaseError = 0;
        public double Channel3AmplitudeError = 0;
        public double Channel3PhaseError = 0;
        public double Channel4AmplitudeError = 0;
        public double Channel4PhaseError = 0;
    }
}
