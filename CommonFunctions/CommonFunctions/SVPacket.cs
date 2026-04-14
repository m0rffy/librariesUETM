using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonFunctions
{
    public class SVPacket
    {
        SVPacketObject[] SVPackets;

        public SVPacketObject SVPacketDecomposition(byte[] Payload)
        {
            string HexPacketsRawString = NumbersConvertion.GetHexStringFrom(Payload);
            string[] HexPacketsRaw = HexPacketsRawString.Split('-');

            string DestinationMAC = HexPacketsRaw[0] + ":" + HexPacketsRaw[1] + ":" + HexPacketsRaw[2] + ":" + HexPacketsRaw[3] + ":" + HexPacketsRaw[4] + ":" + HexPacketsRaw[5];
            string SourceMAC = HexPacketsRaw[6] + ":" + HexPacketsRaw[7] + ":" + HexPacketsRaw[8] + ":" + HexPacketsRaw[9] + ":" + HexPacketsRaw[10] + ":" + HexPacketsRaw[11];
            string EtherType = HexPacketsRaw[12] + HexPacketsRaw[13];
            string APPID = HexPacketsRaw[14] + HexPacketsRaw[15];
            int APPIDInt = NumbersConvertion.HexToInt(APPID);
            string SVPacketLength = HexPacketsRaw[16] + HexPacketsRaw[17];
            int SVPacketLengthInt = NumbersConvertion.HexToInt(SVPacketLength);
            string Reserved1 = HexPacketsRaw[18] + HexPacketsRaw[19];
            string Reserved2 = HexPacketsRaw[20] + HexPacketsRaw[21];
            string savPDUTag = HexPacketsRaw[22] + HexPacketsRaw[23];
            string savPDULength = HexPacketsRaw[24] + HexPacketsRaw[25];
            int savPDULengthInt = NumbersConvertion.HexToInt(savPDULength);
            string NumberOfASDUTag = HexPacketsRaw[26] + HexPacketsRaw[27];
            string NumberOfASDU = HexPacketsRaw[28];
            int NumberOfASDUInt = NumbersConvertion.HexToInt(NumberOfASDU);
            List<string>[] ASDUs = new List<string>[NumberOfASDUInt];

            string SequenceOfASDUTag = HexPacketsRaw[29] + HexPacketsRaw[30];
            string GeneralSequenceOfASDULength = HexPacketsRaw[31] + HexPacketsRaw[32];
            int FirstASDUStartByte = 33;

            for (int k = 0; k < NumberOfASDUInt; k++)
            {
                string SequenceOfASDUStartTag = HexPacketsRaw[FirstASDUStartByte];
                string SequenceOfASDULength = HexPacketsRaw[FirstASDUStartByte + 1];
                int SequenceOfASDU1LengthInt = NumbersConvertion.HexToInt(SequenceOfASDULength) + 2;
                string SequenceOfASDUSvIDTag = HexPacketsRaw[FirstASDUStartByte + 2];
                string SequenceOfASDUSvIDLength = HexPacketsRaw[FirstASDUStartByte + 3];
                int SequenceOfASDU1SvIDLengthInt = NumbersConvertion.HexToInt(SequenceOfASDUSvIDLength);
                string svID = "";
                for (int i = 1; i <= SequenceOfASDU1SvIDLengthInt; i++)
                {
                    svID += HexPacketsRaw[FirstASDUStartByte + 3 + i];
                }
                int ContinueByte1 = FirstASDUStartByte + 3 + SequenceOfASDU1SvIDLengthInt + 1;
                string SmpCntTag = HexPacketsRaw[ContinueByte1] + HexPacketsRaw[ContinueByte1 + 1];
                string SmpCnt = HexPacketsRaw[ContinueByte1 + 2] + HexPacketsRaw[ContinueByte1 + 3];
                string confRevTag = HexPacketsRaw[ContinueByte1 + 4] + HexPacketsRaw[ContinueByte1 + 5];
                string confRev = HexPacketsRaw[ContinueByte1 + 6] + HexPacketsRaw[ContinueByte1 + 7] + HexPacketsRaw[ContinueByte1 + 8] + HexPacketsRaw[ContinueByte1 + 9];
                string smpSynchTag = HexPacketsRaw[ContinueByte1 + 10] + HexPacketsRaw[ContinueByte1 + 11];
                string smpSynch = HexPacketsRaw[ContinueByte1 + 12];
                string SequenceofDataTag = HexPacketsRaw[ContinueByte1 + 13] + HexPacketsRaw[ContinueByte1 + 14];
                string CurrentAmplitude1 = HexPacketsRaw[ContinueByte1 + 15] + HexPacketsRaw[ContinueByte1 + 16] + HexPacketsRaw[ContinueByte1 + 17] + HexPacketsRaw[ContinueByte1 + 18];
                string CurrentQuality1 = HexPacketsRaw[ContinueByte1 + 19] + HexPacketsRaw[ContinueByte1 + 20] + HexPacketsRaw[ContinueByte1 + 21] + HexPacketsRaw[ContinueByte1 + 22];
                string CurrentAmplitude2 = HexPacketsRaw[ContinueByte1 + 23] + HexPacketsRaw[ContinueByte1 + 24] + HexPacketsRaw[ContinueByte1 + 25] + HexPacketsRaw[ContinueByte1 + 26];
                string CurrentQuality2 = HexPacketsRaw[ContinueByte1 + 27] + HexPacketsRaw[ContinueByte1 + 28] + HexPacketsRaw[ContinueByte1 + 29] + HexPacketsRaw[ContinueByte1 + 30];
                string CurrentAmplitude3 = HexPacketsRaw[ContinueByte1 + 31] + HexPacketsRaw[ContinueByte1 + 32] + HexPacketsRaw[ContinueByte1 + 33] + HexPacketsRaw[ContinueByte1 + 34];
                string CurrentQuality3 = HexPacketsRaw[ContinueByte1 + 35] + HexPacketsRaw[ContinueByte1 + 36] + HexPacketsRaw[ContinueByte1 + 37] + HexPacketsRaw[ContinueByte1 + 38];
                string CurrentAmplitude4 = HexPacketsRaw[ContinueByte1 + 39] + HexPacketsRaw[ContinueByte1 + 40] + HexPacketsRaw[ContinueByte1 + 41] + HexPacketsRaw[ContinueByte1 + 42];
                string CurrentQuality4 = HexPacketsRaw[ContinueByte1 + 43] + HexPacketsRaw[ContinueByte1 + 44] + HexPacketsRaw[ContinueByte1 + 45] + HexPacketsRaw[ContinueByte1 + 46];
                string VoltageAmplitude1 = HexPacketsRaw[ContinueByte1 + 47] + HexPacketsRaw[ContinueByte1 + 48] + HexPacketsRaw[ContinueByte1 + 49] + HexPacketsRaw[ContinueByte1 + 50];
                string VoltageQuality1 = HexPacketsRaw[ContinueByte1 + 51] + HexPacketsRaw[ContinueByte1 + 52] + HexPacketsRaw[ContinueByte1 + 53] + HexPacketsRaw[ContinueByte1 + 54];
                string VoltageAmplitude2 = HexPacketsRaw[ContinueByte1 + 55] + HexPacketsRaw[ContinueByte1 + 56] + HexPacketsRaw[ContinueByte1 + 57] + HexPacketsRaw[ContinueByte1 + 58];
                string VoltageQuality2 = HexPacketsRaw[ContinueByte1 + 59] + HexPacketsRaw[ContinueByte1 + 60] + HexPacketsRaw[ContinueByte1 + 61] + HexPacketsRaw[ContinueByte1 + 62];
                string VoltageAmplitude3 = HexPacketsRaw[ContinueByte1 + 63] + HexPacketsRaw[ContinueByte1 + 64] + HexPacketsRaw[ContinueByte1 + 65] + HexPacketsRaw[ContinueByte1 + 66];
                string VoltageQuality3 = HexPacketsRaw[ContinueByte1 + 67] + HexPacketsRaw[ContinueByte1 + 68] + HexPacketsRaw[ContinueByte1 + 69] + HexPacketsRaw[ContinueByte1 + 70];
                string VoltageAmplitude4 = HexPacketsRaw[ContinueByte1 + 71] + HexPacketsRaw[ContinueByte1 + 72] + HexPacketsRaw[ContinueByte1 + 73] + HexPacketsRaw[ContinueByte1 + 74];
                string VoltageQuality4 = HexPacketsRaw[ContinueByte1 + 75] + HexPacketsRaw[ContinueByte1 + 76] + HexPacketsRaw[ContinueByte1 + 77] + HexPacketsRaw[ContinueByte1 + 78];
                ASDUs[k] = new List<string> { svID, SmpCnt, confRev, smpSynch, CurrentAmplitude1, CurrentQuality1, CurrentAmplitude2, CurrentQuality2, CurrentAmplitude3, CurrentQuality3, CurrentAmplitude4, CurrentQuality4, VoltageAmplitude1, VoltageQuality1, VoltageAmplitude2, VoltageQuality2, VoltageAmplitude3, VoltageQuality3, VoltageAmplitude4, VoltageQuality4 };
                FirstASDUStartByte += SequenceOfASDU1LengthInt;
            }
            SVPacketObject SVPacket = new SVPacketObject(DestinationMAC, SourceMAC, APPIDInt, SVPacketLengthInt, savPDULengthInt, NumberOfASDUInt, ASDUs);
            return SVPacket;
        }
    }

    public class SVPacketObject
    {
        public SVPacketObject(string DestinationMAC, string SourceMAC, int APPID, int SVPacketLength, int savPDULength, int NumberOfASDU, List<string>[] ASDUs)
        {
            this.DestinationMAC = DestinationMAC;
            this.SourceMAC = SourceMAC;
            this.APPID = APPID;
            this.SVPacketLength = SVPacketLength;
            this.savPDULength = savPDULength;
            this.NumberOfASDU = NumberOfASDU;
            this.ASDUs = ASDUs;
        }

        public string DestinationMAC = "";
        public string SourceMAC = "";
        public int APPID = 0;
        public int SVPacketLength = 0;
        public int savPDULength = 0;
        public int NumberOfASDU = 0;
        public List<string>[] ASDUs = null;
    }
}