using NModbus;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using static ModBusHelper.ModBusExporterLinker;

//byte = uint8
//ushort = uint16
//uint = uint32
//long = uint64

//htons() — функция, которая превращает hardware to network short

namespace ModBusHelper
{
    public class ModBusProfile
    {
        private ModBusFunctions ModbusFunctionsHelper = new ModBusFunctions();

        // ========================== СТРУКТУРЫ ==========================

        public struct cmns
        {
            public uint SerialNo;
            public uint FmwVer;
            public uint[] SubStNo;
            public byte[] MntPlce;
            public ushort nAchans;
            public ushort nDichan;
            public ushort nDiap;
            public ushort ZdRowInx;
            public ushort eh_cnt_res;
            public ushort dog_cnt_res;
        }

        public struct svs
        {
            public uint cfgRev;
            public dstAddr dstAddr;
            public byte[] sName;
            public ushort nspp;          // кол-во выборок в периоде
            public ushort nasdu;         // кол-во ASDU в SV пакете
            public byte nsig;            // кол-во передаваемых каналов
            public byte nasdu_IU;        // передача показаний тока и напряжения
            public byte gmidentity_ena;  // разрешение добавления идентификатора часов PTP
            public byte smprate_ena;     // разрешение добавления сэмплов в период
        }

        public struct dstAddr
        {
            public byte[] dstMAC;
            public ushort appID;
            public ushort VID;
            public ushort Prio;
        }

        public struct ips
        {
            public byte[] ipMask;
            public byte[] GateWay;
            public byte[] ipAddr;
            public byte[] opts;
            public byte[] name;
        }

        public struct nets
        {
            public byte[] ownAddr;
            public ushort mbport;
            public ips ips;
            public svs svs;
        }

        public struct primct
        {
            public ushort Inom1;
            public ushort Inom2;   // миллиамперы
            public byte[] label;
        }

        public struct meas
        {
            public primct primct;
            public float Rsh;
            public ushort sscle;
            public ushort Nturns;
            public ushort adcrng;
            public byte aStart;
            public byte aStop;
            public uint next_verif;
            public ushort verif_st;
        }

        public struct ptps
        {
            public int acv;
            public uint ofthi;
            public uint oftlo;
            public uint crst;
            public uint aclt;
            public byte[] mmadr;
            public byte[] rqadr;
            public id id;
            public byte rqpd;
            public byte mcai;
        }

        public struct id
        {
            public ushort portNum;
            public byte[] clkId;
        }

        public struct syns
        {
            public int sevOfs;
            public ptps ptps;
            public ushort lmscm;
            public ushort umscm;
            public byte[] res;
        }

        public struct ptpval
        {
            public uint ns;
            public uint slo;
        }

        public struct time
        {
            public ptpval ptpval;
            public ushort ptsecHi;
            public byte[] opts;
        }

        public struct swnf
        {
            public ushort Inn;     // миллиамперы
            public ushort Imax;    // килоамперы
            public byte[] label;
            public byte[] model;
        }

        public struct SCDLY
        {
            public short offd;    // десятые доли миллисекунды
            public short ond;     // десятые доли миллисекунды
        }

        public struct algo
        {
            public ushort Iotc;   // миллиамперы
            public ushort Nn;     // целое
            public float C1;
            public float C2;
            public float C3;
            public float C4;
        }

        public struct contacts
        {
            public ushort invmsk;
            public SCDLY maxdly;
            public SCDLY ajtr;
            public SCDLY[] cdly;
        }

        public struct swrcs
        {
            public ushort swrEna;
            public swnf swnf;
            public algo algo;
            public ushort jrnvol;
            public contacts contacts;
        }

        public struct accor
        {
            public float[] AdcOfs;
            public float[][] actTbl;
            public uint[][] phcTbl;
        }

        public struct GeneralSettings
        {
            public cmns cmns;
            public nets nets;
            public meas meas;
            public syns syns;
            public time time;
            public swrcs swrcs;
            public accor accor;
        }

        // ========================== МЕТОДЫ ЧТЕНИЯ ==========================

        public cmns cmns_Read(IModbusMaster Master)
        {
            ushort addr = 2048;
            ushort mbrsz = 30;
            ushort[] data = Master.ReadHoldingRegisters(0, addr, mbrsz);
            cmns cmnsValues = new cmns();

            cmnsValues.SerialNo = ModbusFunctionsHelper.GetUInt32(data[0], data[1]);
            cmnsValues.FmwVer = ModbusFunctionsHelper.GetUInt32(data[2], data[3]);
            cmnsValues.SubStNo = new uint[2];
            cmnsValues.SubStNo[0] = ModbusFunctionsHelper.GetUInt32(data[4], data[5]);
            cmnsValues.SubStNo[1] = ModbusFunctionsHelper.GetUInt32(data[6], data[7]);

            ushort[] mntRegs = new ushort[16];
            Array.Copy(data, 8, mntRegs, 0, 16);
            cmnsValues.MntPlce = ModbusFunctionsHelper.ConvertUshortArrayToBytes(mntRegs);

            cmnsValues.ZdRowInx = data[0x18];
            cmnsValues.nDiap = data[0x19];
            cmnsValues.nAchans = data[0x1A];
            cmnsValues.nDichan = data[0x1B];
            if (data.Length > 0x1C) cmnsValues.eh_cnt_res = data[0x1C];
            if (data.Length > 0x1D) cmnsValues.dog_cnt_res = data[0x1D];

            return cmnsValues;
        }

        public nets nets_Read(IModbusMaster Master)
        {
            ushort addr = 2304;
            ushort mbrsz = 52;
            ushort[] data = Master.ReadHoldingRegisters(0, addr, mbrsz);
            nets netsValues = new nets();

            // ownAddr (3 регистра)
            ushort[] ownRegs = new ushort[3];
            Array.Copy(data, 0, ownRegs, 0, 3);
            netsValues.ownAddr = ModbusFunctionsHelper.ConvertUshortArrayToBytes(ownRegs);
            netsValues.mbport = data[3];

            // ips (регистры 4-19)
            ushort ipsOffset = 4;
            ushort[] ipMaskRegs = new ushort[2];
            Array.Copy(data, ipsOffset, ipMaskRegs, 0, 2);
            netsValues.ips.ipMask = ModbusFunctionsHelper.ConvertUshortArrayToBytes(ipMaskRegs);
            ushort[] gateRegs = new ushort[2];
            Array.Copy(data, ipsOffset + 2, gateRegs, 0, 2);
            netsValues.ips.GateWay = ModbusFunctionsHelper.ConvertUshortArrayToBytes(gateRegs);
            ushort[] ipAddrRegs = new ushort[2];
            Array.Copy(data, ipsOffset + 4, ipAddrRegs, 0, 2);
            netsValues.ips.ipAddr = ModbusFunctionsHelper.ConvertUshortArrayToBytes(ipAddrRegs);
            ushort[] optsRegs = new ushort[2];
            Array.Copy(data, ipsOffset + 6, optsRegs, 0, 2);
            netsValues.ips.opts = ModbusFunctionsHelper.ConvertUshortArrayToBytes(optsRegs);
            ushort[] nameRegs = new ushort[8];
            Array.Copy(data, ipsOffset + 8, nameRegs, 0, 8);
            netsValues.ips.name = ModbusFunctionsHelper.ConvertUshortArrayToBytes(nameRegs);

            // svs (регистры 20-51)
            ushort svsOffset = 20;
            netsValues.svs.cfgRev = ModbusFunctionsHelper.GetUInt32(data[svsOffset], data[svsOffset + 1]);

            ushort dstOffset = 22;
            ushort[] dstMacRegs = new ushort[3];
            Array.Copy(data, dstOffset, dstMacRegs, 0, 3);
            netsValues.svs.dstAddr.dstMAC = ModbusFunctionsHelper.ConvertUshortArrayToBytes(dstMacRegs);
            netsValues.svs.dstAddr.appID = data[dstOffset + 3];
            netsValues.svs.dstAddr.VID = data[dstOffset + 4];
            netsValues.svs.dstAddr.Prio = data[dstOffset + 5];

            ushort sNameOffset = 28;
            ushort[] sNameRegs = new ushort[18];
            Array.Copy(data, sNameOffset, sNameRegs, 0, 18);
            netsValues.svs.sName = ModbusFunctionsHelper.ConvertUshortArrayToBytes(sNameRegs);

            // Новые поля SV: регистры 2350-2355 (смещения 46-51)
            ushort newFieldsOffset = (ushort)(svsOffset + 0x18); // 20+24=44? wait, correct: 20 + 0x1A = 20+26 = 46
            // Actually: sName occupies 18 registers (28..45), so new fields start at 46.
            newFieldsOffset = (ushort)(sNameOffset + 18); // 28+18 = 46
            netsValues.svs.nspp = data[newFieldsOffset];          // 2350
            netsValues.svs.nasdu = data[newFieldsOffset + 1];     // 2351
            ushort packed = data[newFieldsOffset + 2];            // 2352
            netsValues.svs.nsig = (byte)(packed >> 8);
            netsValues.svs.nasdu_IU = (byte)(packed & 0xFF);
            ushort packed2 = data[newFieldsOffset + 3];           // 2353
            netsValues.svs.gmidentity_ena = (byte)(packed2 >> 8);
            netsValues.svs.smprate_ena = (byte)(packed2 & 0xFF);

            return netsValues;
        }

        public meas meas_Read(IModbusMaster Master)
        {
            ushort addr = 3584;
            ushort mbrsz = 30;
            ushort[] data = Master.ReadHoldingRegisters(0, addr, mbrsz);
            meas measValues = new meas();

            measValues.primct.Inom1 = data[0];          // регистр 3584
            measValues.primct.Inom2 = data[1];          // регистр 3585
            ushort[] labelRegs = new ushort[4];
            Array.Copy(data, 2, labelRegs, 0, 4);
            measValues.primct.label = ModbusFunctionsHelper.ConvertUshortArrayToBytes(labelRegs);
            measValues.Rsh = ModbusFunctionsHelper.GetSingle(data[6], data[7]);
            measValues.sscle = data[8];
            measValues.Nturns = data[9];
            measValues.adcrng = data[10];
            byte[] aBytes = BitConverter.GetBytes(data[11]);
            measValues.aStart = aBytes[0];
            measValues.aStop = aBytes[1];
            measValues.next_verif = ModbusFunctionsHelper.GetUInt32(data[12], data[13]);
            measValues.verif_st = data[14];

            return measValues;
        }

        public syns syns_Read(IModbusMaster Master)
        {
            ushort addr = 3072;
            ushort mbrsz = 28;
            ushort[] data = Master.ReadHoldingRegisters(0, addr, mbrsz);
            syns synsValues = new syns();

            synsValues.sevOfs = ModbusFunctionsHelper.GetInt32(data[0], data[1]);

            ushort ptpsOffset = 2;
            synsValues.ptps.acv = ModbusFunctionsHelper.GetInt32(data[ptpsOffset], data[ptpsOffset + 1]);
            synsValues.ptps.ofthi = ModbusFunctionsHelper.GetUInt32(data[ptpsOffset + 2], data[ptpsOffset + 3]);
            synsValues.ptps.oftlo = ModbusFunctionsHelper.GetUInt32(data[ptpsOffset + 4], data[ptpsOffset + 5]);
            synsValues.ptps.crst = ModbusFunctionsHelper.GetUInt32(data[ptpsOffset + 6], data[ptpsOffset + 7]);
            synsValues.ptps.aclt = ModbusFunctionsHelper.GetUInt32(data[ptpsOffset + 8], data[ptpsOffset + 9]);

            ushort[] mmadrRegs = new ushort[3];
            Array.Copy(data, ptpsOffset + 10, mmadrRegs, 0, 3);
            synsValues.ptps.mmadr = ModbusFunctionsHelper.ConvertUshortArrayToBytes(mmadrRegs);

            ushort[] rqadrRegs = new ushort[3];
            Array.Copy(data, ptpsOffset + 13, rqadrRegs, 0, 3);
            synsValues.ptps.rqadr = ModbusFunctionsHelper.ConvertUshortArrayToBytes(rqadrRegs);

            ushort idOffset = 18;
            synsValues.ptps.id.portNum = data[idOffset];
            ushort[] clkIdRegs = new ushort[4];
            Array.Copy(data, idOffset + 1, clkIdRegs, 0, 4);
            synsValues.ptps.id.clkId = ModbusFunctionsHelper.ConvertUshortArrayToBytes(clkIdRegs);

            byte[] rqpd_mcai = BitConverter.GetBytes(data[ptpsOffset + 21]);
            synsValues.ptps.rqpd = rqpd_mcai[1];
            synsValues.ptps.mcai = rqpd_mcai[0];

            synsValues.lmscm = data[0x18];
            synsValues.umscm = data[0x19];

            return synsValues;
        }

        public time time_Read(IModbusMaster Master)
        {
            ushort addr = 2816;
            ushort mbrsz = 6;
            ushort[] data = Master.ReadHoldingRegisters(0, addr, mbrsz);
            time timeValues = new time();
            timeValues.ptpval.ns = ModbusFunctionsHelper.GetUInt32(data[0], data[1]);
            timeValues.ptpval.slo = ModbusFunctionsHelper.GetUInt32(data[2], data[3]);
            timeValues.ptsecHi = data[4];
            timeValues.opts = BitConverter.GetBytes(data[5]);
            return timeValues;
        }

        public swrcs swrcs_Read(IModbusMaster Master)
        {
            // Сначала читаем cmns, чтобы узнать nDichan
            ushort cmnsAddr = 2048;
            ushort cmnsSize = 30;
            ushort[] cmnsData = Master.ReadHoldingRegisters(0, cmnsAddr, cmnsSize);
            ushort nDichan = cmnsData[0x1B];
            if (nDichan < 3) nDichan = 3; // Минимум 3 для трёхфазной системы

            // Читаем всю группу swrcs (адрес 3328)
            ushort swrcsAddr = 3328;
            ushort swrcsSize = (ushort)(nDichan * 2 + 40);
            ushort[] data = Master.ReadHoldingRegisters(0, swrcsAddr, swrcsSize);

            swrcs swrcsValues = new swrcs();

            // swrEna
            swrcsValues.swrEna = data[0];

            // swnf (адрес 3329)
            ushort swnfBase = 3329;
            swrcsValues.swnf.Inn = data[swnfBase - swrcsAddr];          // регистр 3329
            swrcsValues.swnf.Imax = data[swnfBase - swrcsAddr + 1];    // регистр 3330
                                                                       // label (5 регистров, 3331–3335)
            ushort[] labelRegs = new ushort[5];
            Array.Copy(data, swnfBase - swrcsAddr + 2, labelRegs, 0, 5);
            swrcsValues.swnf.label = ModbusFunctionsHelper.ConvertUshortArrayToBytes(labelRegs);
            // model (16 регистров, 3336–3351)
            ushort[] modelRegs = new ushort[16];
            Array.Copy(data, swnfBase - swrcsAddr + 7, modelRegs, 0, 16);
            swrcsValues.swnf.model = ModbusFunctionsHelper.ConvertUshortArrayToBytes(modelRegs);

            // algo (адрес 3352)
            ushort algoBase = 3352;
            swrcsValues.algo.Iotc = data[algoBase - swrcsAddr];        // регистр 3352
            swrcsValues.algo.Nn = data[algoBase - swrcsAddr + 1];      // регистр 3353
            swrcsValues.algo.C1 = ModbusFunctionsHelper.GetSingle(data[algoBase - swrcsAddr + 2], data[algoBase - swrcsAddr + 3]); // 3354-3355
            swrcsValues.algo.C2 = ModbusFunctionsHelper.GetSingle(data[algoBase - swrcsAddr + 4], data[algoBase - swrcsAddr + 5]); // 3356-3357
            swrcsValues.algo.C3 = ModbusFunctionsHelper.GetSingle(data[algoBase - swrcsAddr + 6], data[algoBase - swrcsAddr + 7]); // 3358-3359
            swrcsValues.algo.C4 = ModbusFunctionsHelper.GetSingle(data[algoBase - swrcsAddr + 8], data[algoBase - swrcsAddr + 9]); // 3360-3361

            // jrnvol (адрес 3362)
            swrcsValues.jrnvol = data[0x22]; // 3362

            // contacts (адрес 3363)
            ushort contactsBase = 3363;
            swrcsValues.contacts.invmsk = data[contactsBase - swrcsAddr];                     // 3363
            swrcsValues.contacts.maxdly.offd = (short)data[contactsBase - swrcsAddr + 1];    // 3364
            swrcsValues.contacts.maxdly.ond = (short)data[contactsBase - swrcsAddr + 2];     // 3365
            swrcsValues.contacts.ajtr.offd = (short)data[contactsBase - swrcsAddr + 3];      // 3366
            swrcsValues.contacts.ajtr.ond = (short)data[contactsBase - swrcsAddr + 4];       // 3367

            // cdly – массив из nDichan записей, каждая по 2 регистра (3368 и далее)
            int cdlyStart = contactsBase - swrcsAddr + 5;
            swrcsValues.contacts.cdly = new SCDLY[nDichan];
            for (int i = 0; i < nDichan; i++)
            {
                swrcsValues.contacts.cdly[i].offd = (short)data[cdlyStart + i * 2];
                swrcsValues.contacts.cdly[i].ond = (short)data[cdlyStart + i * 2 + 1];
            }

            return swrcsValues;
        }

        public accor accor_Read(IModbusMaster Master)
        {
            ushort cmnsAddr = 2048;
            ushort cmnsSize = 30;
            ushort[] cmnsData = Master.ReadHoldingRegisters(0, cmnsAddr, cmnsSize);
            ushort nAchans = cmnsData[0x1A];
            ushort nDiap = cmnsData[0x19];
            if (nAchans < 3) nAchans = 3;
            if (nDiap < 1) nDiap = 1;

            accor accorValues = new accor();

            // AdcOfs
            ushort adcOfsAddr = 3840;
            ushort adcOfsSize = (ushort)(nAchans * 2);
            ushort[] adcData = Master.ReadHoldingRegisters(0, adcOfsAddr, adcOfsSize);
            float[] adcOfs = new float[nAchans];
            for (int i = 0; i < nAchans; i++)
                adcOfs[i] = ModbusFunctionsHelper.GetSingle(adcData[i * 2], adcData[i * 2 + 1]);
            accorValues.AdcOfs = adcOfs;

            // actTbl
            ushort actTblBase = 4096;
            ushort partSize = (ushort)(nDiap * 2);
            List<float> actAll = new List<float>();
            for (int i = 0; i < nAchans; i++)
            {
                ushort addr = (ushort)(actTblBase + partSize * i);
                ushort[] partData = Master.ReadHoldingRegisters(0, addr, partSize);
                for (int j = 0; j < nDiap; j++)
                    actAll.Add(ModbusFunctionsHelper.GetSingle(partData[j * 2], partData[j * 2 + 1]));
            }
            float[][] actTbl = new float[nAchans][];
            for (int i = 0; i < nAchans; i++)
            {
                actTbl[i] = new float[nDiap];
                for (int j = 0; j < nDiap; j++)
                    actTbl[i][j] = actAll[i * nDiap + j];
            }
            accorValues.actTbl = actTbl;

            // phcTbl
            ushort phcTblBase = 5120;
            List<uint> phcAll = new List<uint>();
            for (int i = 0; i < nAchans; i++)
            {
                ushort addr = (ushort)(phcTblBase + partSize * i);
                ushort[] partData = Master.ReadHoldingRegisters(0, addr, partSize);
                for (int j = 0; j < nDiap; j++)
                    phcAll.Add(ModbusFunctionsHelper.GetUInt32(partData[j * 2], partData[j * 2 + 1]));
            }
            uint[][] phcTbl = new uint[nAchans][];
            for (int i = 0; i < nAchans; i++)
            {
                phcTbl[i] = new uint[nDiap];
                for (int j = 0; j < nDiap; j++)
                    phcTbl[i][j] = phcAll[i * nDiap + j];
            }
            accorValues.phcTbl = phcTbl;

            return accorValues;
        }

        // ========================== СТРАНИЦЫ СОСТОЯНИЯ ==========================

        public esp ect_state_page_Read(IModbusMaster Master)
        {
            ushort cmnsAddr = 2048;
            ushort cmnsSize = 30;
            ushort[] cmnsData = Master.ReadHoldingRegisters(0, cmnsAddr, cmnsSize);
            ushort nAchans = cmnsData[0x1A];

            ushort addr = 0x0000;
            ushort mbrsz = (ushort)(2 * nAchans + nAchans / 2 + 2);
            ushort[] data = Master.ReadHoldingRegisters(0, addr, mbrsz);

            esp espValues = new esp();
            ushort[] steRegs = new ushort[2];
            Array.Copy(data, 0, steRegs, 0, 2);
            espValues.ste = ModbusFunctionsHelper.ConvertUshortArrayToBytes(steRegs);

            List<float> rmsList = new List<float>();
            for (int i = 0; i < 2 * nAchans; i += 2)
                rmsList.Add(ModbusFunctionsHelper.GetSingle(data[2 + i], data[2 + i + 1]));
            espValues.rms = rmsList.ToArray();

            ushort diapsOffset = (ushort)(2 * nAchans + 2);
            int diapsCount = nAchans / 2;
            List<ushort> diapsUshort = new List<ushort>();
            for (int i = 0; i < diapsCount; i++)
                diapsUshort.Add(data[diapsOffset + i]);
            espValues.diaps = ModbusFunctionsHelper.ConvertUshortArrayToSbyteArray(diapsUshort.ToArray());
            return espValues;
        }

        public ssp swrct_state_page_Read(IModbusMaster Master)
        {
            ushort cmnsAddr = 2048;
            ushort cmnsSize = 30;
            ushort[] cmnsData = Master.ReadHoldingRegisters(0, cmnsAddr, cmnsSize);
            ushort nDichan = cmnsData[0x1B];

            ushort addr = 0x0100;
            ushort mbrsz = (ushort)(4 * nDichan + 2);
            ushort[] data = Master.ReadHoldingRegisters(0, addr, mbrsz);

            ssp sspValues = new ssp();
            ushort[] steRegs = new ushort[2];
            Array.Copy(data, 0, steRegs, 0, 2);
            sspValues.ste = ModbusFunctionsHelper.ConvertUshortArrayToBytes(steRegs);

            List<SWCNT> cntvList = new List<SWCNT>();
            for (int i = 0; i < 4 * nDichan; i += 4)
            {
                SWCNT sw = new SWCNT();
                sw.Racc = ModbusFunctionsHelper.GetSingle(data[2 + i], data[2 + i + 1]);
                sw.ofcnt = data[2 + i + 2];
                sw.oNacnt = data[2 + i + 3];
                cntvList.Add(sw);
            }
            sspValues.cntv = cntvList.ToArray();
            return sspValues;
        }

        public List<journal_record> journal_record_Read(IModbusMaster Master)
        {
            ushort cmnsAddr = 2048;
            ushort cmnsSize = 30;
            ushort[] cmnsData = Master.ReadHoldingRegisters(0, cmnsAddr, cmnsSize);
            ushort nDichan = cmnsData[0x1B];
            ushort recordLength = (ushort)(4 * nDichan + 10);

            ushort swrcsAddr = 3328;
            ushort swrcsSize = (ushort)(nDichan * 2 + 40);
            ushort[] swrcsData = Master.ReadHoldingRegisters(0, swrcsAddr, swrcsSize);
            ushort jrnvol = swrcsData[0x22];

            List<journal_record> records = new List<journal_record>();
            for (ushort i = 0; i < jrnvol; i++)
            {
                byte[] fileData = Master.ReadFileRecord(0, 1, i, recordLength);
                if (fileData.Length == 0) continue;

                ushort[] regs = ModbusFunctionsHelper.ConvertByteArrayToUShortArray(fileData, true);
                if (regs.Length < 10 + 4 * nDichan) continue;

                journal_record rec = new journal_record();
                rec.hdr.stamp.ns = ModbusFunctionsHelper.GetUInt32(regs[0], regs[1]);
                rec.hdr.stamp.slo = ModbusFunctionsHelper.GetUInt32(regs[2], regs[3]);
                rec.hdr.udt = regs[4];
                byte[] rtypeSub = BitConverter.GetBytes(regs[5]);
                rec.hdr.rtype = rtypeSub[1];
                rec.hdr.subtype = rtypeSub[0];
                rec.Ii = ModbusFunctionsHelper.GetSingle(regs[6], regs[7]);
                rec.Ri = ModbusFunctionsHelper.GetSingle(regs[8], regs[9]);
                rec.rcntv = new SWCNT[nDichan];
                for (int j = 0; j < nDichan; j++)
                {
                    rec.rcntv[j].Racc = ModbusFunctionsHelper.GetSingle(regs[10 + j * 4], regs[10 + j * 4 + 1]);
                    rec.rcntv[j].ofcnt = regs[10 + j * 4 + 2];
                    rec.rcntv[j].oNacnt = regs[10 + j * 4 + 3];
                }
                records.Add(rec);
            }
            return records;
        }

        public GeneralSettings ReadAllSettings(IModbusMaster Master)
        {
            GeneralSettings settings = new GeneralSettings();
            settings.cmns = cmns_Read(Master);
            settings.nets = nets_Read(Master);
            settings.meas = meas_Read(Master);
            settings.syns = syns_Read(Master);
            settings.time = time_Read(Master);
            settings.swrcs = swrcs_Read(Master);
            settings.accor = accor_Read(Master);
            return settings;
        }

        // ========================== МЕТОДЫ ЗАПИСИ ==========================

        public void cmns_Write(IModbusMaster Master, GeneralSettings_TextFormat GeneralSettingsValues_TextFormat, bool AdminRights)
        {
            cmns_TextFormat cmnsValues_TextFormat = GeneralSettingsValues_TextFormat.cmns;

            ushort addr = 2048;

            // SubStNo - Номер подстанции (без изменений)
            ushort SubStNoOffset = 0x04;
            List<ushort> SubStNoList = new List<ushort>();
            foreach (string SubStNo_part_string in cmnsValues_TextFormat.SubStNo)
            {
                uint SubStNo_part_uint = UInt32.Parse(SubStNo_part_string);
                ushort[] SubStNo_part_ushort_array = ModbusFunctionsHelper.ConvertUintToUshortArray(SubStNo_part_uint);
                Array.Reverse(SubStNo_part_ushort_array);
                SubStNoList.AddRange(SubStNo_part_ushort_array);
            }
            ushort[] SubStNo = SubStNoList.ToArray();
            Master.WriteMultipleRegisters(0, (ushort)(addr + SubStNoOffset), SubStNo);

            // MntPlce - Место установки (используем CP1251)
            ushort MntPlceOffset = 0x08;
            ushort[] MntPlce = ModbusFunctionsHelper.ConvertStringToUshortArray(cmnsValues_TextFormat.MntPlce, 32);
            Master.WriteMultipleRegisters(0, (ushort)(addr + MntPlceOffset), MntPlce);

            if (AdminRights == true)
            {
                // SerialNo - Серийный номер устройства (без изменений)
                ushort SerialNoOffset = 0x00;
                ushort[] SerialNo = ModbusFunctionsHelper.ConvertUintToUshortArray(UInt32.Parse(cmnsValues_TextFormat.SerialNo));
                Array.Reverse(SerialNo);
                Master.WriteMultipleRegisters(0, (ushort)(addr + SerialNoOffset), SerialNo);

                // ZdRowInx - Номер строки таблицы коррекции (без изменений)
                ushort ZdRowInxOffset = 0x18;
                ushort ZdRowInx = Convert.ToUInt16(cmnsValues_TextFormat.ZdRowInx);
                ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(addr + ZdRowInxOffset), ZdRowInx);
            }
        }

        public void nets_Write(IModbusMaster Master, GeneralSettings_TextFormat GeneralSettingsValues_TextFormat, bool AdminRights)
        {
            nets_TextFormat netsValues_TextFormat = GeneralSettingsValues_TextFormat.nets;

            ushort addr = 2304;

            // ownAddr (без изменений)
            ushort ownAddrOffset = 0x00;
            ushort[] ownAddr = ModbusFunctionsHelper.ConvertByteArrayToUShortArray(netsValues_TextFormat.ownAddr ?? new byte[6], true);
            Master.WriteMultipleRegisters(0, (ushort)(addr + ownAddrOffset), ownAddr);

            // mbport (без изменений)
            ushort mbportOffset = 0x03;
            ushort mbport = Convert.ToUInt16(netsValues_TextFormat.mbport);
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(addr + mbportOffset), mbport);

            // ips (без изменений)
            ushort ipsaddr = 2308;
            ushort[] ipMask = ModbusFunctionsHelper.ConvertByteArrayToUShortArray(netsValues_TextFormat.ips.ipMask ?? new byte[4], true);
            Master.WriteMultipleRegisters(0, (ushort)(ipsaddr + 0x00), ipMask);
            ushort[] GateWay = ModbusFunctionsHelper.ConvertByteArrayToUShortArray(netsValues_TextFormat.ips.GateWay ?? new byte[4], true);
            Master.WriteMultipleRegisters(0, (ushort)(ipsaddr + 0x02), GateWay);
            ushort[] ipAddr = ModbusFunctionsHelper.ConvertByteArrayToUShortArray(netsValues_TextFormat.ips.ipAddr ?? new byte[4], true);
            Master.WriteMultipleRegisters(0, (ushort)(ipsaddr + 0x04), ipAddr);
            ushort[] opts = ModbusFunctionsHelper.ConvertByteArrayToUShortArray(netsValues_TextFormat.ips.opts ?? new byte[4], true);
            Master.WriteMultipleRegisters(0, (ushort)(ipsaddr + 0x06), opts);
            // name – используем CP1251
            ushort[] name = ModbusFunctionsHelper.ConvertStringToUshortArray(netsValues_TextFormat.ips.name ?? "", 16);
            Master.WriteMultipleRegisters(0, (ushort)(ipsaddr + 0x08), name);

            // svs
            ushort svsaddr = 2324;
            ushort[] cfgRev = ModbusFunctionsHelper.ConvertUintToUshortArray(UInt32.Parse(netsValues_TextFormat.svs.cfgRev ?? "0"));
            Array.Reverse(cfgRev);
            Master.WriteMultipleRegisters(0, (ushort)(svsaddr + 0x00), cfgRev);

            ushort dstAddraddr = 2326;
            ushort[] dstMAC = ModbusFunctionsHelper.ConvertByteArrayToUShortArray(netsValues_TextFormat.svs.dstAddr.dstMAC ?? new byte[6], true);
            Master.WriteMultipleRegisters(0, (ushort)(dstAddraddr + 0x00), dstMAC);
            ushort appID = Convert.ToUInt16(netsValues_TextFormat.svs.dstAddr.appID ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(dstAddraddr + 0x03), appID);
            ushort VID = Convert.ToUInt16(netsValues_TextFormat.svs.dstAddr.VID ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(dstAddraddr + 0x04), VID);
            ushort Prio = Convert.ToUInt16(netsValues_TextFormat.svs.dstAddr.Prio ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(dstAddraddr + 0x05), Prio);

            // sName – используем CP1251
            ushort sNameOffset = 0x08;
            ushort[] sName = ModbusFunctionsHelper.ConvertStringToUshortArray(netsValues_TextFormat.svs.sName ?? "", 36);
            Master.WriteMultipleRegisters(0, (ushort)(svsaddr + sNameOffset), sName);

            // Новые поля SV (без изменений)
            ushort newFieldsAddr = (ushort)(svsaddr + 0x1A);
            ushort nspp = Convert.ToUInt16(netsValues_TextFormat.svs.nspp ?? "0");
            ushort nasdu = Convert.ToUInt16(netsValues_TextFormat.svs.nasdu ?? "0");
            byte nsig = Convert.ToByte(netsValues_TextFormat.svs.nsig ?? "0");
            byte nasdu_IU = Convert.ToByte(netsValues_TextFormat.svs.nasdu_IU ?? "0");
            byte gmidentity_ena = Convert.ToByte(netsValues_TextFormat.svs.gmidentity_ena ?? "0");
            byte smprate_ena = Convert.ToByte(netsValues_TextFormat.svs.smprate_ena ?? "0");
            ushort packed = (ushort)((nsig << 8) | nasdu_IU);
            ushort packed2 = (ushort)((gmidentity_ena << 8) | smprate_ena);
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(newFieldsAddr), nspp);
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(newFieldsAddr + 1), nasdu);
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(newFieldsAddr + 2), packed);
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(newFieldsAddr + 3), packed2);
        }

        // Исправленный метод meas_Write в ModBusProfile.cs
        public void meas_Write(IModbusMaster Master, GeneralSettings_TextFormat GeneralSettingsValues_TextFormat, bool AdminRights)
        {
            meas_TextFormat meas = GeneralSettingsValues_TextFormat.meas;
            ushort addr = 3584;
            ushort primctaddr = 3584;

            // Inom1
            ushort Inom1 = Convert.ToUInt16(meas.primct.Inom1 ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, primctaddr, Inom1);

            // Inom2
            ushort Inom2 = Convert.ToUInt16(meas.primct.Inom2 ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(primctaddr + 1), Inom2);

            // label – используем CP1251
            ushort[] label = ModbusFunctionsHelper.ConvertStringToUshortArray(meas.primct.label ?? "", 8);
            Master.WriteMultipleRegisters(0, (ushort)(primctaddr + 2), label);

            // sscle
            ushort sscle = Convert.ToUInt16(meas.sscle ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(addr + 8), sscle);

            // aStart / aStop
            byte aStart = Convert.ToByte(meas.aStart ?? "0");
            byte aStop = Convert.ToByte(meas.aStop ?? "0");
            ushort ADCNumbers = BitConverter.ToUInt16(new byte[2] { aStart, aStop }, 0);
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(addr + 11), ADCNumbers);

            if (AdminRights)
            {
                // Rsh
                float rshVal = float.Parse(meas.Rsh ?? "0", CultureInfo.InvariantCulture);
                ushort[] Rsh = ModbusFunctionsHelper.ConvertFloatToUshortArray(rshVal);
                Array.Reverse(Rsh);
                Master.WriteMultipleRegisters(0, (ushort)(addr + 6), Rsh);

                // Nturns
                ushort Nturns = Convert.ToUInt16(meas.Nturns ?? "0");
                ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(addr + 9), Nturns);

                // adcrng
                ushort adcrng = Convert.ToUInt16(meas.adcrng ?? "0");
                ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(addr + 10), adcrng);
            }
        }

        public void syns_Write(IModbusMaster Master, GeneralSettings_TextFormat GeneralSettingsValues_TextFormat, bool AdminRights)
        {
            syns_TextFormat synsValues_TextFormat = GeneralSettingsValues_TextFormat.syns;

            ushort addr = 3072;

            // sevOfs
            ushort[] sevOfs = ModbusFunctionsHelper.ConvertIntToUshortArray(Int32.Parse(synsValues_TextFormat.sevOfs ?? "0"));
            Array.Reverse(sevOfs);
            Master.WriteMultipleRegisters(0, (ushort)(addr + 0x00), sevOfs);

            ushort ptpsaddr = 3074;

            // acv
            ushort[] acv = ModbusFunctionsHelper.ConvertIntToUshortArray(Int32.Parse(synsValues_TextFormat.ptps.acv ?? "0"));
            Array.Reverse(acv);
            Master.WriteMultipleRegisters(0, (ushort)(ptpsaddr + 0x00), acv);

            // ofthi
            ushort[] ofthi = ModbusFunctionsHelper.ConvertUintToUshortArray(UInt32.Parse(synsValues_TextFormat.ptps.ofthi ?? "0"));
            Array.Reverse(ofthi);
            Master.WriteMultipleRegisters(0, (ushort)(ptpsaddr + 0x02), ofthi);

            // oftlo
            ushort[] oftlo = ModbusFunctionsHelper.ConvertUintToUshortArray(UInt32.Parse(synsValues_TextFormat.ptps.oftlo ?? "0"));
            Array.Reverse(oftlo);
            Master.WriteMultipleRegisters(0, (ushort)(ptpsaddr + 0x04), oftlo);

            // crst
            ushort[] crst = ModbusFunctionsHelper.ConvertUintToUshortArray(UInt32.Parse(synsValues_TextFormat.ptps.crst ?? "0"));
            Array.Reverse(crst);
            Master.WriteMultipleRegisters(0, (ushort)(ptpsaddr + 0x06), crst);

            // aclt
            ushort[] aclt = ModbusFunctionsHelper.ConvertUintToUshortArray(UInt32.Parse(synsValues_TextFormat.ptps.aclt ?? "0"));
            Array.Reverse(aclt);
            Master.WriteMultipleRegisters(0, (ushort)(ptpsaddr + 0x08), aclt);

            // mmadr
            ushort[] mmadr = ModbusFunctionsHelper.ConvertByteArrayToUShortArray(synsValues_TextFormat.ptps.mmadr ?? new byte[6], true);
            Master.WriteMultipleRegisters(0, (ushort)(ptpsaddr + 0x0A), mmadr);

            // rqadr
            ushort[] rqadr = ModbusFunctionsHelper.ConvertByteArrayToUShortArray(synsValues_TextFormat.ptps.rqadr ?? new byte[6], true);
            Master.WriteMultipleRegisters(0, (ushort)(ptpsaddr + 0x0D), rqadr);

            // id
            ushort idaddr = 3090;
            ushort portNum = Convert.ToUInt16(synsValues_TextFormat.ptps.id.portNum ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(idaddr + 0x00), portNum);
            // clkId – используем CP1251
            ushort[] clkId = ModbusFunctionsHelper.ConvertStringToUshortArray(synsValues_TextFormat.ptps.id.clkId ?? "", 8);
            Master.WriteMultipleRegisters(0, (ushort)(idaddr + 0x01), clkId);

            // rqpd / mcai
            byte rqpd = Convert.ToByte(synsValues_TextFormat.ptps.rqpd ?? "0");
            byte mcai = Convert.ToByte(synsValues_TextFormat.ptps.mcai ?? "0");
            ushort SyncNumbers = BitConverter.ToUInt16(new byte[2] { rqpd, mcai }, 0);
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(ptpsaddr + 0x15), SyncNumbers);

            // lmscm
            ushort lmscm = Convert.ToUInt16(synsValues_TextFormat.lmscm ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(addr + 0x18), lmscm);

            // umscm
            ushort umscm = Convert.ToUInt16(synsValues_TextFormat.umscm ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(addr + 0x19), umscm);
        }

        public void time_Write(IModbusMaster Master, GeneralSettings_TextFormat GeneralSettingsValues_TextFormat, bool AdminRights)
        {
            time_TextFormat timeValues_TextFormat = GeneralSettingsValues_TextFormat.time;

            ushort addr = 2816;
            ushort ptpvaladdr = 2816;

            ushort[] ns = ModbusFunctionsHelper.ConvertUintToUshortArray(UInt32.Parse(timeValues_TextFormat.ptpval.ns ?? "0"));
            Array.Reverse(ns);
            Master.WriteMultipleRegisters(0, (ushort)(ptpvaladdr + 0x00), ns);

            ushort[] slo = ModbusFunctionsHelper.ConvertUintToUshortArray(UInt32.Parse(timeValues_TextFormat.ptpval.slo ?? "0"));
            Array.Reverse(slo);
            Master.WriteMultipleRegisters(0, (ushort)(ptpvaladdr + 0x02), slo);

            ushort ptsecHi = Convert.ToUInt16(timeValues_TextFormat.ptsecHi ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(addr + 0x04), ptsecHi);

            ushort optsOffset = 0x05;
            byte[] optsBytes = timeValues_TextFormat.opts ?? new byte[2];
            ushort[] opts = ModbusFunctionsHelper.ConvertByteArrayToUShortArray(optsBytes, true);
            Master.WriteMultipleRegisters(0, (ushort)(addr + optsOffset), opts);
        }

        // Исправленный метод swrcs_Write в ModBusProfile.cs
        public void swrcs_Write(IModbusMaster Master, GeneralSettings_TextFormat GeneralSettingsValues_TextFormat, bool AdminRights)
        {
            swrcs_TextFormat swrcs = GeneralSettingsValues_TextFormat.swrcs;
            ushort num = 3328;

            // swrEna
            ushort value = Convert.ToUInt16(swrcs.swrEna ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, num, value);

            // swnf
            ushort num2 = 3329;
            ushort num3 = 0;
            if (!ushort.TryParse(swrcs.swnf.Inn, out num3))
                num3 = Convert.ToUInt16(swrcs.swnf.Inn ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, num2, num3);

            ushort value2 = Convert.ToUInt16(swrcs.swnf.Imax ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(num2 + 1), value2);

            // label – используем CP1251
            ushort[] data = ModbusFunctionsHelper.ConvertStringToUshortArray(swrcs.swnf.label ?? "", 10);
            Master.WriteMultipleRegisters(0, (ushort)(num2 + 2), data);

            // model – используем CP1251
            ushort[] data2 = ModbusFunctionsHelper.ConvertStringToUshortArray(swrcs.swnf.model ?? "", 32);
            Master.WriteMultipleRegisters(0, (ushort)(num2 + 7), data2);

            // algo
            ushort num4 = 3352;
            ushort num5 = 0;
            if (!ushort.TryParse(swrcs.algo.Iotc, out num5))
                num5 = Convert.ToUInt16(swrcs.algo.Iotc ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, num4, num5);

            ushort value3 = Convert.ToUInt16(swrcs.algo.Nn ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, (ushort)(num4 + 1), value3);

            float number = float.Parse(swrcs.algo.C1 ?? "0", CultureInfo.InvariantCulture);
            ushort[] array = ModbusFunctionsHelper.ConvertFloatToUshortArray(number);
            Array.Reverse(array);
            Master.WriteMultipleRegisters(0, (ushort)(num4 + 2), array);

            float number2 = float.Parse(swrcs.algo.C2 ?? "0", CultureInfo.InvariantCulture);
            ushort[] array2 = ModbusFunctionsHelper.ConvertFloatToUshortArray(number2);
            Array.Reverse(array2);
            Master.WriteMultipleRegisters(0, (ushort)(num4 + 4), array2);

            float number3 = float.Parse(swrcs.algo.C3 ?? "0", CultureInfo.InvariantCulture);
            ushort[] array3 = ModbusFunctionsHelper.ConvertFloatToUshortArray(number3);
            Array.Reverse(array3);
            Master.WriteMultipleRegisters(0, (ushort)(num4 + 6), array3);

            float number4 = float.Parse(swrcs.algo.C4 ?? "0", CultureInfo.InvariantCulture);
            ushort[] array4 = ModbusFunctionsHelper.ConvertFloatToUshortArray(number4);
            Array.Reverse(array4);
            Master.WriteMultipleRegisters(0, (ushort)(num4 + 8), array4);

            // contacts
            ushort num6 = 3363;
            ushort value4 = Convert.ToUInt16(swrcs.contacts.invmsk ?? "0");
            ModbusFunctionsHelper.WriteSingleRegisterThroughLimitation(Master, num6, value4);

            short num7 = Convert.ToInt16(swrcs.contacts.ajtr.offd ?? "0");
            short num8 = Convert.ToInt16(swrcs.contacts.ajtr.ond ?? "0");
            Master.WriteMultipleRegisters(0, (ushort)(num6 + 3), new ushort[2]
            {
                ModbusFunctionsHelper.ConvertShortToUshort(num7),
                ModbusFunctionsHelper.ConvertShortToUshort(num8)
            });

            ushort num9 = 0;
            if (!string.IsNullOrEmpty(GeneralSettingsValues_TextFormat.cmns.nDichan))
            {
                num9 = Convert.ToUInt16(GeneralSettingsValues_TextFormat.cmns.nDichan);
                for (int i = 0; i < num9 && i < swrcs.contacts.cdly.Length; i++)
                {
                    short num10 = swrcs.contacts.cdly[i].offd;
                    short num11 = swrcs.contacts.cdly[i].ond;
                    Master.WriteMultipleRegisters(0, (ushort)(num6 + 5 + i * 2), new ushort[2]
                    {
                        ModbusFunctionsHelper.ConvertShortToUshort(num10),
                        ModbusFunctionsHelper.ConvertShortToUshort(num11)
                    });
                }
            }
        }

        public void accor_Write(IModbusMaster Master, GeneralSettings_TextFormat GeneralSettingsValues_TextFormat, bool AdminRights)
        {
            // Без изменений (нет строк)
            if (AdminRights == true)
            {
                ushort nAchans = 0, nDiap = 0;
                if (!string.IsNullOrEmpty(GeneralSettingsValues_TextFormat.cmns.nAchans) && !string.IsNullOrEmpty(GeneralSettingsValues_TextFormat.cmns.nDiap))
                {
                    nAchans = Convert.ToUInt16(GeneralSettingsValues_TextFormat.cmns.nAchans);
                    nDiap = Convert.ToUInt16(GeneralSettingsValues_TextFormat.cmns.nDiap);
                }
                else return;

                accor_TextFormat accorValues_TextFormat = GeneralSettingsValues_TextFormat.accor;

                // AdcOfs
                ushort AdcOfsaddr = 3840;
                List<ushort> AdcOfs_List = new List<ushort>();
                foreach (string AdcOfs_parameter in accorValues_TextFormat.AdcOfs)
                {
                    ushort[] AdcOfs_parameter_ushorts = ModbusFunctionsHelper.ConvertTextFloatToUshortArray(AdcOfs_parameter);
                    Array.Reverse(AdcOfs_parameter_ushorts);
                    AdcOfs_List.AddRange(AdcOfs_parameter_ushorts);
                }
                Master.WriteMultipleRegisters(0, AdcOfsaddr, AdcOfs_List.ToArray());

                // actTbl
                List<List<ushort>> actTbl_List = new List<List<ushort>>();
                for (int i = 0; i < nAchans; i++)
                {
                    List<ushort> actTbl_ExactChannel = new List<ushort>();
                    for (int j = 0; j < nDiap; j++)
                    {
                        ushort[] ExactactTbl = ModbusFunctionsHelper.ConvertFloatToUshortArray(accorValues_TextFormat.actTbl[i][j]);
                        Array.Reverse(ExactactTbl);
                        actTbl_ExactChannel.AddRange(ExactactTbl);
                    }
                    actTbl_List.Add(actTbl_ExactChannel);
                }
                ushort[][] actTblArrayofArrays = actTbl_List.Select(l => l.ToArray()).ToArray();
                int GeneralLength = nAchans * nDiap;
                ushort[] actTbl_part1 = new ushort[GeneralLength];
                actTblArrayofArrays[0].CopyTo(actTbl_part1, 0);
                actTblArrayofArrays[1].CopyTo(actTbl_part1, actTblArrayofArrays[0].Length);
                ushort[] actTbl_part2 = new ushort[GeneralLength];
                actTblArrayofArrays[2].CopyTo(actTbl_part2, 0);
                actTblArrayofArrays[3].CopyTo(actTbl_part2, actTblArrayofArrays[2].Length);

                ushort actTbladdr = 4096;
                Master.WriteMultipleRegisters(0, actTbladdr, actTbl_part1);
                Master.WriteMultipleRegisters(0, (ushort)(actTbladdr + actTbl_part1.Length), actTbl_part2);

                // phcTbl
                List<List<ushort>> phcTbl_List = new List<List<ushort>>();
                for (int i = 0; i < nAchans; i++)
                {
                    List<ushort> phcTbl_ExactChannel = new List<ushort>();
                    for (int j = 0; j < nDiap; j++)
                    {
                        ushort[] ExactphcTbl = ModbusFunctionsHelper.ConvertUintToUshortArray(accorValues_TextFormat.phcTbl[i][j]);
                        Array.Reverse(ExactphcTbl);
                        phcTbl_ExactChannel.AddRange(ExactphcTbl);
                    }
                    phcTbl_List.Add(phcTbl_ExactChannel);
                }
                ushort[][] phcTblArrayofArrays = phcTbl_List.Select(l => l.ToArray()).ToArray();
                ushort[] phcTbl_part1 = new ushort[GeneralLength];
                phcTblArrayofArrays[0].CopyTo(phcTbl_part1, 0);
                phcTblArrayofArrays[1].CopyTo(phcTbl_part1, phcTblArrayofArrays[0].Length);
                ushort[] phcTbl_part2 = new ushort[GeneralLength];
                phcTblArrayofArrays[2].CopyTo(phcTbl_part2, 0);
                phcTblArrayofArrays[3].CopyTo(phcTbl_part2, phcTblArrayofArrays[2].Length);

                ushort phcTbladdr = 5120;
                Master.WriteMultipleRegisters(0, phcTbladdr, phcTbl_part1);
                Master.WriteMultipleRegisters(0, (ushort)(phcTbladdr + phcTbl_part1.Length), phcTbl_part2);
            }
        }

        // ========================== ДОПОЛНИТЕЛЬНЫЕ СТРУКТУРЫ ==========================

        public struct esp
        {
            public byte[] ste;
            public float[] rms;
            public sbyte[] diaps;
        }

        public struct SWCNT
        {
            public float Racc;
            public ushort ofcnt;
            public ushort oNacnt;
        }

        public struct ssp
        {
            public byte[] ste;
            public SWCNT[] cntv;
        }

        public struct RECHDR
        {
            public ptpval stamp;
            public ushort udt;
            public byte rtype;
            public byte subtype;
        }

        public struct journal_record
        {
            public RECHDR hdr;
            public float Ii;
            public float Ri;
            public SWCNT[] rcntv;
        }
    }
}