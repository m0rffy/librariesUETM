using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonFunctions
{
    public class Calculation
    {
        public AnalyticsObject GetRMSandPhaseErrorFromSVPackets(SVPacketObject[] SVPacketsArray, string EtalonSVStreamName, string TestDeviceSVStreamName)
        {
            int EtalonASDUsCounter = 0;
            int TestingDeviceASDUsCounter = 0;

            foreach (SVPacketObject SVPacket in SVPacketsArray)
            {
                foreach (List<string> ASDU in SVPacket.ASDUs)
                {
                    string StreamName = NumbersConvertion.HexToText(ASDU[0]);
                    if (StreamName == EtalonSVStreamName)
                    {
                        EtalonASDUsCounter++;
                    }
                    else if (StreamName == TestDeviceSVStreamName)
                    {
                        TestingDeviceASDUsCounter++;
                    }
                }
            }

            SampleObject[] EtalonSamplesChannel1 = new SampleObject[EtalonASDUsCounter];
            SampleObject[] TestingDeviceSamplesChannel1 = new SampleObject[TestingDeviceASDUsCounter];
            SampleObject[] TestingDeviceSamplesChannel2 = new SampleObject[TestingDeviceASDUsCounter];
            SampleObject[] TestingDeviceSamplesChannel3 = new SampleObject[TestingDeviceASDUsCounter];
            SampleObject[] TestingDeviceSamplesChannel4 = new SampleObject[TestingDeviceASDUsCounter];

            int EtalonSamplesCounter = 0;
            int TestingDeviceSamplesCounter = 0;

            foreach (SVPacketObject SVPacket in SVPacketsArray)
            {
                foreach (List<string> ASDU in SVPacket.ASDUs)
                {
                    if (NumbersConvertion.HexToText(ASDU[0]) == EtalonSVStreamName)
                    {
                        EtalonSamplesChannel1[EtalonSamplesCounter] = new SampleObject(NumbersConvertion.HexToInt(ASDU[1]), NumbersConvertion.HexToLong(ASDU[4]), NumbersConvertion.HexToInt(ASDU[3]), ASDU[0]);
                        EtalonSamplesCounter++;
                    }

                    else if (NumbersConvertion.HexToText(ASDU[0]) == TestDeviceSVStreamName)
                    {
                        TestingDeviceSamplesChannel1[TestingDeviceSamplesCounter] = new SampleObject(NumbersConvertion.HexToInt(ASDU[1]), NumbersConvertion.HexToLong(ASDU[4]), NumbersConvertion.HexToInt(ASDU[3]), ASDU[0]);
                        TestingDeviceSamplesChannel2[TestingDeviceSamplesCounter] = new SampleObject(NumbersConvertion.HexToInt(ASDU[1]), NumbersConvertion.HexToLong(ASDU[6]), NumbersConvertion.HexToInt(ASDU[3]), ASDU[0]);
                        TestingDeviceSamplesChannel3[TestingDeviceSamplesCounter] = new SampleObject(NumbersConvertion.HexToInt(ASDU[1]), NumbersConvertion.HexToLong(ASDU[8]), NumbersConvertion.HexToInt(ASDU[3]), ASDU[0]);
                        TestingDeviceSamplesChannel4[TestingDeviceSamplesCounter] = new SampleObject(NumbersConvertion.HexToInt(ASDU[1]), NumbersConvertion.HexToLong(ASDU[10]), NumbersConvertion.HexToInt(ASDU[3]), ASDU[0]);
                        TestingDeviceSamplesCounter++;
                    }
                }
            }

            List<SampleObject> FirstDeviceChannel1Samples = new List<SampleObject>();
            List<SampleObject> SecondDeviceChannel1Samples = new List<SampleObject>();
            List<SampleObject> SecondDeviceChannel2Samples = new List<SampleObject>();
            List<SampleObject> SecondDeviceChannel3Samples = new List<SampleObject>();
            List<SampleObject> SecondDeviceChannel4Samples = new List<SampleObject>();

            foreach (SampleObject Sample in EtalonSamplesChannel1)
            {
                FirstDeviceChannel1Samples.Add(Sample);
            }

            foreach (SampleObject Sample in TestingDeviceSamplesChannel1)
            {
                SecondDeviceChannel1Samples.Add(Sample);
            }

            foreach (SampleObject Sample in TestingDeviceSamplesChannel2)
            {
                SecondDeviceChannel2Samples.Add(Sample);
            }

            foreach (SampleObject Sample in TestingDeviceSamplesChannel3)
            {
                SecondDeviceChannel3Samples.Add(Sample);
            }

            foreach (SampleObject Sample in TestingDeviceSamplesChannel4)
            {
                SecondDeviceChannel4Samples.Add(Sample);
            }

            int EtalonAntiZeroCounter = 0;
            int TestingDeviceChannel1AntiZeroCounter = 0;
            int TestingDeviceChannel2AntiZeroCounter = 0;
            int TestingDeviceChannel3AntiZeroCounter = 0;
            int TestingDeviceChannel4AntiZeroCounter = 0;
            long EtalonFirstSampleAmplitude = FindFirstNonNullAmplitude(FirstDeviceChannel1Samples, EtalonAntiZeroCounter);
            long TestingDeviceChannel1FirstSampleAmplitude = FindFirstNonNullAmplitude(SecondDeviceChannel1Samples, TestingDeviceChannel1AntiZeroCounter);
            long TestingDeviceChannel2FirstSampleAmplitude = FindFirstNonNullAmplitude(SecondDeviceChannel2Samples, TestingDeviceChannel2AntiZeroCounter);
            long TestingDeviceChannel3FirstSampleAmplitude = FindFirstNonNullAmplitude(SecondDeviceChannel3Samples, TestingDeviceChannel3AntiZeroCounter);
            long TestingDeviceChannel4FirstSampleAmplitude = FindFirstNonNullAmplitude(SecondDeviceChannel4Samples, TestingDeviceChannel4AntiZeroCounter);

            // Если первая выборка эталона отрицательная
            if (EtalonFirstSampleAmplitude < 0)
            {
                int Channel1TestFlag = 0;
                int Channel2TestFlag = 0;
                int Channel3TestFlag = 0;
                int Channel4TestFlag = 0;

                if (TestingDeviceChannel1FirstSampleAmplitude < 0)
                {
                    Channel1TestFlag = 1;
                }

                if (TestingDeviceChannel2FirstSampleAmplitude < 0)
                {
                    Channel2TestFlag = 1;
                }

                if (TestingDeviceChannel3FirstSampleAmplitude < 0)
                {
                    Channel3TestFlag = 1;
                }

                if (TestingDeviceChannel4FirstSampleAmplitude < 0)
                {
                    Channel4TestFlag = 1;
                }

                // Если в эталоне ток идёт на увеличение (ближайший переход от минуса к плюсу)
                if (EtalonFirstSampleAmplitude < FirstDeviceChannel1Samples[EtalonAntiZeroCounter + 1].Amplitude)
                {
                    // То все каналы которые также как и эталон находятся в минусе выводим в плюс

                    while (FirstDeviceChannel1Samples[EtalonAntiZeroCounter].Amplitude < 0)
                    {
                        EtalonAntiZeroCounter++;
                    }

                    if (FirstDeviceChannel1Samples[EtalonAntiZeroCounter].Amplitude == 0)
                    {
                        EtalonAntiZeroCounter++;
                    }

                    if (Channel1TestFlag == 1)
                    {
                        while (SecondDeviceChannel1Samples[TestingDeviceChannel1AntiZeroCounter].Amplitude < 0)
                        {
                            TestingDeviceChannel1AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel1Samples[TestingDeviceChannel1AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel1AntiZeroCounter++;
                        }
                    }

                    if (Channel2TestFlag == 1)
                    {
                        while (SecondDeviceChannel2Samples[TestingDeviceChannel2AntiZeroCounter].Amplitude < 0)
                        {
                            TestingDeviceChannel2AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel2Samples[TestingDeviceChannel2AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel2AntiZeroCounter++;
                        }
                    }

                    if (Channel3TestFlag == 1)
                    {
                        while (SecondDeviceChannel3Samples[TestingDeviceChannel3AntiZeroCounter].Amplitude < 0)
                        {
                            TestingDeviceChannel3AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel3Samples[TestingDeviceChannel3AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel3AntiZeroCounter++;
                        }
                    }

                    if (Channel4TestFlag == 1)
                    {
                        while (SecondDeviceChannel4Samples[TestingDeviceChannel4AntiZeroCounter].Amplitude < 0)
                        {
                            TestingDeviceChannel4AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel4Samples[TestingDeviceChannel4AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel4AntiZeroCounter++;
                        }
                    }
                }

                // Если в эталоне ток идёт на уменьшение (ближайший переход от плюса к минусу)
                else
                {
                    // То все каналы которые находятся в плюсе выводим на минус
                    if (Channel1TestFlag == 0)
                    {
                        while (SecondDeviceChannel1Samples[TestingDeviceChannel1AntiZeroCounter].Amplitude > 0)
                        {
                            TestingDeviceChannel1AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel1Samples[TestingDeviceChannel1AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel1AntiZeroCounter++;
                        }
                    }

                    if (Channel2TestFlag == 0)
                    {
                        while (SecondDeviceChannel2Samples[TestingDeviceChannel2AntiZeroCounter].Amplitude > 0)
                        {
                            TestingDeviceChannel2AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel2Samples[TestingDeviceChannel2AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel2AntiZeroCounter++;
                        }
                    }

                    if (Channel3TestFlag == 0)
                    {
                        while (SecondDeviceChannel3Samples[TestingDeviceChannel3AntiZeroCounter].Amplitude > 0)
                        {
                            TestingDeviceChannel3AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel3Samples[TestingDeviceChannel3AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel3AntiZeroCounter++;
                        }
                    }

                    if (Channel4TestFlag == 0)
                    {
                        while (SecondDeviceChannel4Samples[TestingDeviceChannel4AntiZeroCounter].Amplitude > 0)
                        {
                            TestingDeviceChannel4AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel4Samples[TestingDeviceChannel4AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel4AntiZeroCounter++;
                        }
                    }
                }
            }

            // Если первая выборка в эталоне положительная
            else if (EtalonFirstSampleAmplitude > 0)
            {
                int Channel1TestFlag = 0;
                int Channel2TestFlag = 0;
                int Channel3TestFlag = 0;
                int Channel4TestFlag = 0;

                if (TestingDeviceChannel1FirstSampleAmplitude > 0)
                {
                    Channel1TestFlag = 1;
                }

                if (TestingDeviceChannel2FirstSampleAmplitude > 0)
                {
                    Channel2TestFlag = 1;
                }

                if (TestingDeviceChannel3FirstSampleAmplitude > 0)
                {
                    Channel3TestFlag = 1;
                }

                if (TestingDeviceChannel4FirstSampleAmplitude > 0)
                {
                    Channel4TestFlag = 1;
                }

                // Если в эталоне ток идёт на увеличение (ближайший переход от минуса к плюсу)
                if (EtalonFirstSampleAmplitude < FirstDeviceChannel1Samples[EtalonAntiZeroCounter + 1].Amplitude)
                {
                    // То все каналы которые находятся в минусе выводим в плюс
                    if (Channel1TestFlag == 0)
                    {
                        while (SecondDeviceChannel1Samples[TestingDeviceChannel1AntiZeroCounter].Amplitude < 0)
                        {
                            TestingDeviceChannel1AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel1Samples[TestingDeviceChannel1AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel1AntiZeroCounter++;
                        }
                    }

                    if (Channel2TestFlag == 0)
                    {
                        while (SecondDeviceChannel2Samples[TestingDeviceChannel2AntiZeroCounter].Amplitude < 0)
                        {
                            TestingDeviceChannel2AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel2Samples[TestingDeviceChannel2AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel2AntiZeroCounter++;
                        }
                    }

                    if (Channel3TestFlag == 0)
                    {
                        while (SecondDeviceChannel3Samples[TestingDeviceChannel3AntiZeroCounter].Amplitude < 0)
                        {
                            TestingDeviceChannel3AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel3Samples[TestingDeviceChannel3AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel3AntiZeroCounter++;
                        }
                    }

                    if (Channel4TestFlag == 0)
                    {
                        while (SecondDeviceChannel4Samples[TestingDeviceChannel4AntiZeroCounter].Amplitude < 0)
                        {
                            TestingDeviceChannel4AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel4Samples[TestingDeviceChannel4AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel4AntiZeroCounter++;
                        }
                    }
                }

                // Если в эталоне ток идёт на уменьшение (ближайший переход от плюса к минусу)
                else
                {
                    // То все каналы которые также как и эталон находятся в плюсе выводим на минус

                    while (FirstDeviceChannel1Samples[EtalonAntiZeroCounter].Amplitude > 0)
                    {
                        EtalonAntiZeroCounter++;
                    }

                    if (FirstDeviceChannel1Samples[EtalonAntiZeroCounter].Amplitude == 0)
                    {
                        EtalonAntiZeroCounter++;
                    }

                    if (Channel1TestFlag == 1)
                    {
                        while (SecondDeviceChannel1Samples[TestingDeviceChannel1AntiZeroCounter].Amplitude > 0)
                        {
                            TestingDeviceChannel1AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel1Samples[TestingDeviceChannel1AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel1AntiZeroCounter++;
                        }
                    }

                    if (Channel2TestFlag == 1)
                    {
                        while (SecondDeviceChannel2Samples[TestingDeviceChannel2AntiZeroCounter].Amplitude > 0)
                        {
                            TestingDeviceChannel2AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel2Samples[TestingDeviceChannel2AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel2AntiZeroCounter++;
                        }
                    }

                    if (Channel3TestFlag == 1)
                    {
                        while (SecondDeviceChannel3Samples[TestingDeviceChannel3AntiZeroCounter].Amplitude > 0)
                        {
                            TestingDeviceChannel3AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel3Samples[TestingDeviceChannel3AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel3AntiZeroCounter++;
                        }
                    }

                    if (Channel4TestFlag == 1)
                    {
                        while (SecondDeviceChannel4Samples[TestingDeviceChannel4AntiZeroCounter].Amplitude > 0)
                        {
                            TestingDeviceChannel4AntiZeroCounter++;
                        }

                        if (SecondDeviceChannel4Samples[TestingDeviceChannel4AntiZeroCounter].Amplitude == 0)
                        {
                            TestingDeviceChannel4AntiZeroCounter++;
                        }
                    }
                }
            }

            FirstDeviceChannel1Samples.RemoveRange(0, Math.Min(EtalonAntiZeroCounter, FirstDeviceChannel1Samples.Count));
            SecondDeviceChannel1Samples.RemoveRange(0, Math.Min(TestingDeviceChannel1AntiZeroCounter, SecondDeviceChannel1Samples.Count));
            SecondDeviceChannel2Samples.RemoveRange(0, Math.Min(TestingDeviceChannel2AntiZeroCounter, SecondDeviceChannel2Samples.Count));
            SecondDeviceChannel3Samples.RemoveRange(0, Math.Min(TestingDeviceChannel3AntiZeroCounter, SecondDeviceChannel3Samples.Count));
            SecondDeviceChannel4Samples.RemoveRange(0, Math.Min(TestingDeviceChannel4AntiZeroCounter, SecondDeviceChannel4Samples.Count));

            double[] PrimaryValuesChannel1 = GetPrimaryValues(FirstDeviceChannel1Samples, SecondDeviceChannel1Samples);
            double[] PrimaryValuesChannel2 = GetPrimaryValues(FirstDeviceChannel1Samples, SecondDeviceChannel2Samples);
            double[] PrimaryValuesChannel3 = GetPrimaryValues(FirstDeviceChannel1Samples, SecondDeviceChannel3Samples);
            double[] PrimaryValuesChannel4 = GetPrimaryValues(FirstDeviceChannel1Samples, SecondDeviceChannel4Samples);

            AnalyticsObject AnalyticsObject = new AnalyticsObject(PrimaryValuesChannel1[0], PrimaryValuesChannel1[2], PrimaryValuesChannel1[3], PrimaryValuesChannel2[2], PrimaryValuesChannel2[3], PrimaryValuesChannel3[2], PrimaryValuesChannel3[3], PrimaryValuesChannel4[2], PrimaryValuesChannel4[3]);

            return AnalyticsObject;
        }

        private int FindSyncRange(List<SampleObject> Device1Samples, List<SampleObject> Device2Samples)
        {
            int m = 0;
            int flag = 0;
            int FirstSample = 0;
            while (flag == 0)
            {
                for (int p = m; p < Device1Samples.Count; p++)
                {
                    if ((Device1Samples[p].SynchStatus == 1 || Device1Samples[p].SynchStatus == 2) && (Device2Samples[p].SynchStatus == 1 || Device2Samples[p].SynchStatus == 2))
                    {
                        FirstSample = p;
                        break;
                    }
                    else
                    {
                        continue;
                    }
                }

                for (int k = 0; k < 3000; k++)
                {
                    if ((Device1Samples[FirstSample + k].SynchStatus == 1 || Device1Samples[FirstSample + k].SynchStatus == 2) && (Device2Samples[FirstSample + k].SynchStatus == 1 || Device2Samples[FirstSample + k].SynchStatus == 2))
                    {
                        flag = 1;
                        continue;
                    }
                    else
                    {
                        m = FirstSample + k;
                        flag = 0;
                        break;
                    }
                }
            }
            return FirstSample;
        }

        private List<double> GetZeroCrossArray(List<SampleObject> DeviceSamples, int FirstSample)
        {
            int WorkSamplesCount = DeviceSamples.Count - FirstSample;
            double[] ResultTime = new double[WorkSamplesCount];
            long[] OriginalSin = new long[WorkSamplesCount];
            double TIMEPERIOD = 1.0 / 4800.0;
            List<double> ZeroCrossArray = new List<double>();
            int PlusMinusFlag = 0;
            int MinusPlusFlag = 0;

            // Если первая выборка имеет отрицательный знак
            if (DeviceSamples[FirstSample].Amplitude < 0)
            {
                MinusPlusFlag = 1;

                // ищем первый переход через 0
                while (DeviceSamples[FirstSample].Amplitude < 0)
                {
                    FirstSample++;
                }

                // если есть точка 0
                if (DeviceSamples[FirstSample].Amplitude == 0)
                {
                    ZeroCrossArray.Add(TIMEPERIOD * DeviceSamples[FirstSample].SampleCount);

                    for (int m = 0; m < WorkSamplesCount - FirstSample; m++)
                    {
                        OriginalSin[m] = DeviceSamples[FirstSample + m].Amplitude;
                    }

                    for (int u = 0; u < WorkSamplesCount - FirstSample; u++)
                    {
                        ResultTime[u] = u * TIMEPERIOD;
                    }

                    int blink = 40;
                    int Repeats = WorkSamplesCount / 48;
                    if (Repeats > 5)
                    {
                        Repeats = Repeats - 5;
                    }

                    for (int n = 0; n < Repeats; n++)
                    {
                        if (MinusPlusFlag == 1)
                        {
                            // ищем переход через 0 с плюса на минус
                            while (OriginalSin[blink] > 0)
                            {
                                blink++;
                            }

                            if (OriginalSin[blink] == 0)
                            {
                                ZeroCrossArray.Add(TIMEPERIOD * DeviceSamples[blink + FirstSample].SampleCount);
                            }

                            else
                            {
                                double time = GetTimeOfZeroCross(OriginalSin[blink - 1], OriginalSin[blink], TIMEPERIOD);
                                ZeroCrossArray.Add(time + (TIMEPERIOD * DeviceSamples[blink + FirstSample - 1].SampleCount));
                            }

                            blink = blink + 40;

                            MinusPlusFlag = 0;
                            PlusMinusFlag = 1;
                        }

                        else if (PlusMinusFlag == 1)
                        {
                            // ищем переход через 0 с минуса на плюс
                            while (OriginalSin[blink] > 0)
                            {
                                blink++;
                            }

                            if (OriginalSin[blink] == 0)
                            {
                                ZeroCrossArray.Add(TIMEPERIOD * DeviceSamples[blink + FirstSample].SampleCount);
                            }

                            else
                            {
                                double time = GetTimeOfZeroCross(OriginalSin[blink - 1], OriginalSin[blink], TIMEPERIOD);
                                ZeroCrossArray.Add(time + (TIMEPERIOD * DeviceSamples[blink + FirstSample - 1].SampleCount));
                            }

                            blink = blink + 40;

                            MinusPlusFlag = 1;
                            PlusMinusFlag = 0;
                        }

                    }
                }

                // Если точки 0 нету
                else
                {
                    FirstSample = FirstSample - 1;

                    for (int m = 0; m < WorkSamplesCount - FirstSample; m++)
                    {
                        OriginalSin[m] = DeviceSamples[FirstSample + m].Amplitude;
                    }

                    for (int u = 0; u < WorkSamplesCount - FirstSample; u++)
                    {
                        ResultTime[u] = u * TIMEPERIOD;
                    }

                    double time1 = GetTimeOfZeroCross(OriginalSin[0], OriginalSin[1], TIMEPERIOD);
                    ZeroCrossArray.Add(time1 + (TIMEPERIOD * DeviceSamples[FirstSample].SampleCount));

                    int blink = 40;
                    int Repeats = WorkSamplesCount / 48;
                    if (Repeats > 5)
                    {
                        Repeats = Repeats - 5;
                    }

                    for (int n = 0; n < Repeats; n++)
                    {
                        if (MinusPlusFlag == 1)
                        {
                            // ищем переход через 0 с плюса на минус
                            while (OriginalSin[blink] > 0)
                            {
                                blink++;
                            }

                            if (OriginalSin[blink] == 0)
                            {
                                ZeroCrossArray.Add(TIMEPERIOD * DeviceSamples[blink + FirstSample].SampleCount);
                            }

                            else
                            {
                                double time = GetTimeOfZeroCross(OriginalSin[blink - 1], OriginalSin[blink], TIMEPERIOD);
                                ZeroCrossArray.Add(time + (TIMEPERIOD * DeviceSamples[blink + FirstSample - 1].SampleCount));
                            }

                            blink = blink + 40;

                            MinusPlusFlag = 0;
                            PlusMinusFlag = 1;
                        }

                        else if (PlusMinusFlag == 1)
                        {
                            // ищем переход через 0 с минуса на плюс
                            while (OriginalSin[blink] > 0)
                            {
                                blink++;
                            }

                            if (OriginalSin[blink] == 0)
                            {
                                ZeroCrossArray.Add(TIMEPERIOD * DeviceSamples[blink + FirstSample].SampleCount);
                            }

                            else
                            {
                                double time = GetTimeOfZeroCross(OriginalSin[blink - 1], OriginalSin[blink], TIMEPERIOD);
                                ZeroCrossArray.Add(time + (TIMEPERIOD * DeviceSamples[blink + FirstSample - 1].SampleCount));
                            }

                            blink = blink + 40;

                            MinusPlusFlag = 1;
                            PlusMinusFlag = 0;
                        }
                    }
                }

                return ZeroCrossArray;
            }

            else // если первая выборка из результирующего массива имеет положительный знак
            {
                PlusMinusFlag = 1;

                // ищем первый переход через 0 с положительной полуволны
                while (DeviceSamples[FirstSample].Amplitude > 0)
                {
                    FirstSample++;
                }

                // если есть точка 0
                if (DeviceSamples[FirstSample].Amplitude == 0)
                {
                    ZeroCrossArray.Add(TIMEPERIOD * DeviceSamples[FirstSample].SampleCount);

                    for (int m = 0; m < WorkSamplesCount - FirstSample; m++)
                    {
                        OriginalSin[m] = DeviceSamples[FirstSample + m].Amplitude;
                    }

                    for (int u = 0; u < WorkSamplesCount - FirstSample; u++)
                    {
                        ResultTime[u] = u * TIMEPERIOD;
                    }

                    int blink = 40;
                    int Repeats = WorkSamplesCount / 48;
                    if (Repeats > 5)
                    {
                        Repeats = Repeats - 5;
                    }

                    for (int n = 0; n < Repeats; n++)
                    {
                        if (MinusPlusFlag == 1)
                        {
                            // ищем переход через 0 с плюса на минус
                            while (OriginalSin[blink] > 0)
                            {
                                blink++;
                            }

                            if (OriginalSin[blink] == 0)
                            {
                                ZeroCrossArray.Add(TIMEPERIOD * DeviceSamples[blink + FirstSample].SampleCount);
                            }

                            else
                            {
                                double time = GetTimeOfZeroCross(OriginalSin[blink - 1], OriginalSin[blink], TIMEPERIOD);
                                ZeroCrossArray.Add(time + (TIMEPERIOD * DeviceSamples[blink + FirstSample - 1].SampleCount));
                            }

                            blink = blink + 40;

                            MinusPlusFlag = 0;
                            PlusMinusFlag = 1;
                        }

                        else if (PlusMinusFlag == 1)
                        {
                            // ищем переход через 0 с минуса на плюс
                            while (OriginalSin[blink] < 0)
                            {
                                blink++;
                            }

                            if (OriginalSin[blink] == 0)
                            {
                                ZeroCrossArray.Add(TIMEPERIOD * DeviceSamples[blink + FirstSample].SampleCount);
                            }

                            else
                            {
                                double time = GetTimeOfZeroCross(OriginalSin[blink - 1], OriginalSin[blink], TIMEPERIOD);
                                ZeroCrossArray.Add(time + (TIMEPERIOD * DeviceSamples[blink + FirstSample - 1].SampleCount));
                            }

                            blink = blink + 40;

                            MinusPlusFlag = 1;
                            PlusMinusFlag = 0;
                        }

                    }
                }

                // Если точки 0 нету
                else
                {
                    FirstSample = FirstSample - 1;

                    for (int m = 0; m < WorkSamplesCount - FirstSample; m++)
                    {
                        OriginalSin[m] = DeviceSamples[FirstSample + m].Amplitude;
                    }

                    for (int u = 0; u < WorkSamplesCount - FirstSample; u++)
                    {
                        ResultTime[u] = u * TIMEPERIOD;
                    }

                    double time1 = GetTimeOfZeroCross(OriginalSin[0], OriginalSin[1], TIMEPERIOD);
                    ZeroCrossArray.Add(time1 + (TIMEPERIOD * DeviceSamples[FirstSample].SampleCount));

                    int blink = 40;
                    int Repeats = WorkSamplesCount / 48;
                    if (Repeats > 5)
                    {
                        Repeats = Repeats - 5;
                    }

                    for (int n = 0; n < Repeats; n++)
                    {
                        if (MinusPlusFlag == 1)
                        {
                            // ищем переход через 0 с плюса на минус
                            while (OriginalSin[blink] > 0)
                            {
                                blink++;
                            }

                            if (OriginalSin[blink] == 0)
                            {
                                ZeroCrossArray.Add(TIMEPERIOD * DeviceSamples[blink + FirstSample].SampleCount);
                            }

                            else
                            {
                                double time = GetTimeOfZeroCross(OriginalSin[blink - 1], OriginalSin[blink], TIMEPERIOD);
                                ZeroCrossArray.Add(time + (TIMEPERIOD * DeviceSamples[blink + FirstSample - 1].SampleCount));
                            }

                            blink = blink + 40;

                            MinusPlusFlag = 0;
                            PlusMinusFlag = 1;
                        }

                        else if (PlusMinusFlag == 1)
                        {
                            // ищем переход через 0 с минуса на плюс
                            while (OriginalSin[blink] < 0)
                            {
                                blink++;
                            }

                            if (OriginalSin[blink] == 0)
                            {
                                ZeroCrossArray.Add(TIMEPERIOD * DeviceSamples[blink + FirstSample].SampleCount);
                            }

                            else
                            {
                                double time = GetTimeOfZeroCross(OriginalSin[blink - 1], OriginalSin[blink], TIMEPERIOD);
                                ZeroCrossArray.Add(time + (TIMEPERIOD * DeviceSamples[blink + FirstSample - 1].SampleCount));
                            }

                            blink = blink + 40;

                            MinusPlusFlag = 1;
                            PlusMinusFlag = 0;
                        }
                    }
                }

                return ZeroCrossArray;
            }
        }

        private double[] GetRMSValueSingleDevice(List<SampleObject> DeviceSamples, int FirstSample)
        {
            double[] ResultTime = new double[482];
            double[] OriginalSin = new double[482];
            double TIMEPERIOD = 1.0 / 4800.0;
            double[] Results = new double[2];
            double time1 = 0;
            double time2 = 0;

            if (DeviceSamples[FirstSample].Amplitude < 0)
            {
                // ищем первый переход через 0
                while (DeviceSamples[FirstSample].Amplitude < 0)
                {
                    FirstSample++;
                }

                if (DeviceSamples[FirstSample].Amplitude == 0)
                {
                    time1 = 0;

                    for (int m = 0; m < 482; m++)
                    {
                        OriginalSin[m] = DeviceSamples[FirstSample + m].Amplitude;
                    }

                    for (int u = 0; u < 482; u++)
                    {
                        ResultTime[u] = u * TIMEPERIOD;
                    }

                    int v = 90;

                    // ищем второй переход oчерез 0
                    while (OriginalSin[v] < 0)
                    {
                        v++;
                    }

                    if (OriginalSin[v] == 0)
                    {
                        time2 = ResultTime[v];
                    }

                    else
                    {
                        double a2 = (OriginalSin[v - 1] - OriginalSin[v]) / (ResultTime[v - 1] - ResultTime[v]);

                        double b2 = OriginalSin[v - 1] - (ResultTime[v - 1] * a2);

                        time2 = -b2 / a2;
                    }

                    double difference = time2 - time1;
                    double frequency = 1 / difference;

                    // Вычисляем погрешность
                    double Accumulator = 0;

                    for (int r = 0; r < 482; r++)
                    {
                        Accumulator += (OriginalSin[r] * OriginalSin[r]) * ((1.0 / frequency) / 482);
                    }
                    double RMSValue = Math.Sqrt(Accumulator / (1.0 / frequency));

                    Results[0] = RMSValue;
                    Results[1] = TIMEPERIOD * DeviceSamples[FirstSample].SampleCount;
                }

                else
                {
                    FirstSample = FirstSample - 1;

                    for (int m = 0; m < 482; m++)
                    {
                        OriginalSin[m] = DeviceSamples[FirstSample + m].Amplitude;
                    }

                    for (int u = 0; u < 482; u++)
                    {
                        ResultTime[u] = u * TIMEPERIOD;
                    }

                    double a1 = (OriginalSin[0] - OriginalSin[1]) / (ResultTime[0] - ResultTime[1]);

                    double b1 = OriginalSin[0] - (ResultTime[0] * a1);

                    time1 = -b1 / a1;

                    double x0 = time1;

                    double sec = ResultTime[0];

                    double rad = sec * 2 * Math.PI / 0.02;

                    double delta = x0 - ResultTime[0];

                    double x0_rad = rad - delta * 2 * Math.PI / 0.02;

                    int v = 90;

                    // ищем второй переход через 0
                    while (OriginalSin[v] < 0)
                    {
                        v++;
                    }

                    double a2 = (OriginalSin[v - 1] - OriginalSin[v]) / (ResultTime[v - 1] - ResultTime[v]);

                    double b2 = OriginalSin[v - 1] - (ResultTime[v - 1] * a2);

                    time2 = -b2 / a2;

                    double difference = time2 - time1;
                    double frequency = 1 / difference;

                    // Вычисляем погрешность
                    double Accumulator = 0;

                    for (int r = 0; r < 482; r++)
                    {
                        Accumulator += (OriginalSin[r] * OriginalSin[r]) * ((1.0 / frequency) / 482);
                    }
                    double RMSValue = Math.Sqrt(Accumulator / (1.0 / frequency));

                    Results[0] = RMSValue;
                    Results[1] = time1 + (TIMEPERIOD * DeviceSamples[FirstSample].SampleCount);
                }

                return Results;
            }

            else // если первая выборка из результирующего массива имеет положительный знак
            {
                // ищем первый переход через 0 с положительной полуволны
                while (DeviceSamples[FirstSample].Amplitude > 0)
                {
                    FirstSample++;
                }

                if (DeviceSamples[FirstSample].Amplitude == 0)
                {
                    FirstSample++;
                }

                // ищем первый переход через 0 с отрицательной полуволны
                while (DeviceSamples[FirstSample].Amplitude < 0)
                {
                    FirstSample++;
                }

                if (DeviceSamples[FirstSample].Amplitude == 0)
                {
                    time1 = 0;

                    for (int m = 0; m < 482; m++)
                    {
                        OriginalSin[m] = DeviceSamples[FirstSample + m].Amplitude;
                    }

                    for (int u = 0; u < 482; u++)
                    {
                        ResultTime[u] = u * TIMEPERIOD;
                    }

                    int v = 90;

                    // ищем второй переход через 0
                    while (OriginalSin[v] < 0)
                    {
                        v++;
                    }

                    if (OriginalSin[v] == 0)
                    {
                        time2 = ResultTime[v];
                    }

                    else
                    {
                        double a2 = (OriginalSin[v - 1] - OriginalSin[v]) / (ResultTime[v - 1] - ResultTime[v]);
                        double b2 = OriginalSin[v - 1] - (ResultTime[v - 1] * a2);
                        time2 = -b2 / a2;
                    }

                    double difference = time2 - time1;
                    double frequency = 1 / difference;

                    // Вычисляем погрешность
                    double Accumulator = 0;

                    for (int r = 0; r < 482; r++)
                    {
                        Accumulator += (OriginalSin[r] * OriginalSin[r]) * ((1.0 / frequency) / 482);
                    }
                    double RMSValue = Math.Sqrt(Accumulator / (1.0 / frequency));

                    Results[0] = RMSValue;
                    Results[1] = TIMEPERIOD * DeviceSamples[FirstSample].SampleCount;
                }

                else
                {
                    FirstSample = FirstSample - 1;

                    for (int m = 0; m < 482; m++)
                    {
                        OriginalSin[m] = DeviceSamples[FirstSample + m].Amplitude;
                    }

                    for (int u = 0; u < 482; u++)
                    {
                        ResultTime[u] = u * TIMEPERIOD;
                    }

                    double a1 = (OriginalSin[0] - OriginalSin[1]) / (ResultTime[0] - ResultTime[1]);

                    double b1 = OriginalSin[0] - (ResultTime[0] * a1);

                    time1 = -b1 / a1;

                    double x0 = time1;

                    double sec = ResultTime[0];

                    double rad = sec * 2 * Math.PI / 0.02;

                    double delta = x0 - ResultTime[0];

                    double x0_rad = rad - delta * 2 * Math.PI / 0.02;

                    int v = 90;

                    // ищем второй переход через 0
                    while (OriginalSin[v] < 0)
                    {
                        v++;
                    }

                    double a2 = (OriginalSin[v - 1] - OriginalSin[v]) / (ResultTime[v - 1] - ResultTime[v]);

                    double b2 = OriginalSin[v - 1] - (ResultTime[v - 1] * a2);

                    time2 = -b2 / a2;

                    double difference = time2 - time1;
                    double frequency = 1 / difference;

                    // Вычисляем погрешность
                    double Accumulator = 0;

                    for (int r = 0; r < 482; r++)
                    {
                        Accumulator += (OriginalSin[r] * OriginalSin[r]) * ((1.0 / frequency) / 482);
                    }
                    double RMSValue = Math.Sqrt(Accumulator / (1.0 / frequency));

                    Results[0] = RMSValue;
                    Results[1] = time1 + (TIMEPERIOD * DeviceSamples[FirstSample].SampleCount);
                }
                return Results;
            }
        }

        private double[] GetPrimaryValues(List<SampleObject> Device1Samples, List<SampleObject> Device2Samples)
        {
            int FirstSampleSync = FindSyncRange(Device1Samples, Device2Samples);
            double[] Device1Results = GetRMSValueSingleDevice(Device1Samples, FirstSampleSync);
            double[] Device2Results = GetRMSValueSingleDevice(Device2Samples, FirstSampleSync);
            double RMSError = ((Device1Results[0] - Device2Results[0]) / Device1Results[0]) * 100;
            List<double> Device1ZeroCrossArray = GetZeroCrossArray(Device1Samples, FirstSampleSync);
            List<double> Device2ZeroCrossArray = GetZeroCrossArray(Device2Samples, FirstSampleSync);
            int Device1ZeroCrossCount = Device1ZeroCrossArray.Count;
            int Device2ZeroCrossCount = Device2ZeroCrossArray.Count;
            int MinimumCount = 0;
            if (Device1ZeroCrossCount < Device2ZeroCrossCount)
            {
                MinimumCount = Device1ZeroCrossCount;
            }
            else
            {
                MinimumCount = Device2ZeroCrossCount;
            }
            List<double> PhaseErrorArray = new List<double>();
            for (int i = 0; i < MinimumCount; i++)
            {
                PhaseErrorArray.Add(Device1ZeroCrossArray[i] - Device2ZeroCrossArray[i]);
            }
            double MaxZeroCross = PhaseErrorArray.Max();
            double MinZeroCross = PhaseErrorArray.Min();
            double PhaseErrorAccumulator = 0;
            for (int i = 0; i < PhaseErrorArray.Count; i++)
            {
                PhaseErrorAccumulator += PhaseErrorArray[i];
            }
            double PhaseErrorSeconds = PhaseErrorAccumulator / PhaseErrorArray.Count;
            double PhaseErrorMinutes = 18 * PhaseErrorSeconds * 1000 * 60;
            double[] Values = new double[5];
            Values[0] = Device1Results[0];
            Values[1] = Device2Results[0];
            Values[2] = RMSError;
            Values[3] = PhaseErrorSeconds;
            Values[4] = PhaseErrorMinutes;
            return Values;
        }

        private long FindFirstNonNullAmplitude(List<SampleObject> DeviceSamples, int AntiZeroCounter)
        {
            long FirstNonNullSampleAmplitude = 0;

            if (DeviceSamples[0].Amplitude != 0)
            {
                FirstNonNullSampleAmplitude = DeviceSamples[0].Amplitude;
            }
            else
            {
                while (DeviceSamples[AntiZeroCounter].Amplitude == 0)
                {
                    AntiZeroCounter++;
                }

                FirstNonNullSampleAmplitude = DeviceSamples[AntiZeroCounter].Amplitude;
            }
            return FirstNonNullSampleAmplitude;
        }

        private double GetTimeOfZeroCross(double x0, double x1, double TIMEPERIOD)
        {
            double a = (x0 - x1) / TIMEPERIOD;

            double time = Math.Abs(x0 / a);
            return time;
        }
    }
}
