using NModbus;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Sockets;
using static ModBusHelper.ModBusExporterLinker;
using static ModBusHelper.ModBusProfile;


namespace ModBusHelper
{
    public class ModBusExporterLinker
    {
        private ModBusFunctions ModbusFunctionsHelper = new ModBusFunctions();
        private ModBusProfile ModBusProfileHelper = new ModBusProfile();
        public bool WasRebootCommandSent { get; private set; }

        // ========================== ТЕКСТОВЫЕ СТРУКТУРЫ ==========================

        public struct cmns_TextFormat
        {
            public string SerialNo;
            public string FmwVer;
            public string[] SubStNo;
            public string MntPlce;
            public string nAchans;
            public string nDichan;
            public string nDiap;
            public string ZdRowInx;
            public string eh_cnt_res;
            public string dog_cnt_res;
        }

        public struct svs_TextFormat
        {
            public string cfgRev;
            public dstAddr_TextFormat dstAddr;
            public string sName;
            public string nspp;
            public string nasdu;
            public string nsig;
            public string nasdu_IU;
            public string gmidentity_ena;
            public string smprate_ena;
        }

        public struct dstAddr_TextFormat
        {
            public byte[] dstMAC;
            public string appID;
            public string VID;
            public string Prio;
        }

        public struct ips_TextFormat
        {
            public byte[] ipMask;
            public byte[] GateWay;
            public byte[] ipAddr;
            public byte[] opts;
            public string name;
        }

        public struct nets_TextFormat
        {
            public byte[] ownAddr;
            public string mbport;
            public ips_TextFormat ips;
            public svs_TextFormat svs;
            public gss_TextFormat gss;
            public bool gssSupported;
        }

        public struct primct_TextFormat
        {
            public string Inom1;
            public string Inom2;
            public string label;
        }

        public struct verif_TextFormat
        {
            public string next_verif;
            public string verif_st;
        }
        public struct qb_TextFormat
        {
            public string method_oor;
            public string out_of_range;
            public string delta_od;
            public string repl_smps_od;
            public string min_lev_od;
            public string delta_osc;
            public string repl_smps_osc;
            public string min_lev_osc;
            public string imbal_lev;
            public string qb_mod;
            public string[] qb_msk;  // массив строк
        }

        public struct meas_TextFormat
        {
            public primct_TextFormat primct;
            public string Rsh;
            public string sscle;
            public string Nturns;
            public string adcrng;
            public string aStart;
            public string aStop;
            public string adc_osf;               
            public verif_TextFormat verif;     
            public qb_TextFormat qb;
        }

        public struct id_TextFormat
        {
            public string portNum;
            public string clkId;
        }

        public struct ptps_TextFormat
        {
            public string acv;
            public string ofthi;
            public string oftlo;
            public string crst;
            public string aclt;
            public byte[] mmadr;
            public byte[] rqadr;
            public id_TextFormat id;
            public string rqpd;
            public string mcai;
        }

        public struct syns_TextFormat
        {
            public string sevOfs;
            public ptps_TextFormat ptps;
            public string lmscm;
            public string umscm;
            public string[] res;
        }

        public struct ptpval_TextFormat
        {
            public string ns;
            public string slo;
        }

        public struct time_TextFormat
        {
            public ptpval_TextFormat ptpval;
            public string ptsecHi;
            public byte[] opts;
        }

        public struct SCDLY_TextFormat
        {
            public string offd;
            public string ond;
        }

        public struct SCDLY_cdly_TextFormat
        {
            public short offd;
            public short ond;
        }

        public struct swnf_TextFormat
        {
            public string Inn;
            public string Imax;
            public string label;
            public string model;
        }

        public struct algo_TextFormat
        {
            public string Iotc;
            public string Nn;
            public string C1;
            public string C2;
            public string C3;
            public string C4;
        }

        public struct contacts_TextFormat
        {
            public string invmsk;
            public SCDLY_TextFormat maxdly;
            public SCDLY_TextFormat ajtr;
            public SCDLY_cdly_TextFormat[] cdly;
        }

        public struct swrcs_TextFormat
        {
            public string swrEna;
            public swnf_TextFormat swnf;
            public algo_TextFormat algo;
            public string jrnvol;
            public contacts_TextFormat contacts;
        }

        public struct accor_TextFormat
        {
            public string[] AdcOfs;
            public float[][] actTbl;
            public uint[][] phcTbl;
        }

        public struct GeneralSettings_TextFormat
        {
            public cmns_TextFormat cmns;
            public nets_TextFormat nets;
            public meas_TextFormat meas;
            public syns_TextFormat syns;
            public time_TextFormat time;
            public swrcs_TextFormat swrcs;
            public accor_TextFormat accor;
        }

        public struct gss_dstAddr_TextFormat
        {
            public byte[] dstMAC;
            public string appID;
            public string VID;
            public string Prio;
        }

        public struct gss_TextFormat
        {
            public string cfgRev;
            public gss_dstAddr_TextFormat dstAddr;
            public string sml;
            public string brmode;
        }


        // ========================== МЕТОДЫ ==========================


        public static string FormatFloatNoExponent(float value) {
            return value.ToString("0.0###############", CultureInfo.InvariantCulture);
        }
        public GeneralSettings ReadSettings(Tuple<TcpClient, IModbusMaster> connection)
        {
            return ModBusProfileHelper.ReadAllSettings(connection.Item2);
        }

        public GeneralSettings_TextFormat get_GeneralSettings_Text(Tuple<TcpClient, IModbusMaster> connection)
        {
            GeneralSettings raw = ReadSettings(connection);
            GeneralSettings_TextFormat txt = new GeneralSettings_TextFormat();

            // cmns
            txt.cmns.SerialNo = raw.cmns.SerialNo.ToString() ?? "";
            txt.cmns.FmwVer = raw.cmns.FmwVer.ToString() ?? "";
            if (raw.cmns.SubStNo != null)
            {
                txt.cmns.SubStNo = new string[raw.cmns.SubStNo.Length];
                for (int i = 0; i < raw.cmns.SubStNo.Length; i++)
                    txt.cmns.SubStNo[i] = raw.cmns.SubStNo[i].ToString() ?? "";
            }
            else txt.cmns.SubStNo = new string[0];

            // MntPlce – Чтение в CP1251
            if (raw.cmns.MntPlce != null)
                txt.cmns.MntPlce = ModbusFunctionsHelper.ConvertByteArrayToString(raw.cmns.MntPlce).Replace("\0", "");
            else txt.cmns.MntPlce = "";

            txt.cmns.nAchans = raw.cmns.nAchans.ToString() ?? "";
            txt.cmns.nDichan = raw.cmns.nDichan.ToString() ?? "";
            txt.cmns.nDiap = raw.cmns.nDiap.ToString() ?? "";
            txt.cmns.ZdRowInx = raw.cmns.ZdRowInx.ToString() ?? "";
            txt.cmns.eh_cnt_res = raw.cmns.eh_cnt_res.ToString() ?? "0";
            txt.cmns.dog_cnt_res = raw.cmns.dog_cnt_res.ToString() ?? "0";

            // nets
            txt.nets.ownAddr = raw.nets.ownAddr ?? new byte[0];
            txt.nets.mbport = raw.nets.mbport.ToString() ?? "";
            txt.nets.svs.cfgRev = raw.nets.svs.cfgRev.ToString() ?? "";
            txt.nets.svs.dstAddr.dstMAC = raw.nets.svs.dstAddr.dstMAC ?? new byte[0];
            txt.nets.svs.dstAddr.appID = raw.nets.svs.dstAddr.appID.ToString() ?? "";
            txt.nets.svs.dstAddr.VID = raw.nets.svs.dstAddr.VID.ToString() ?? "";
            txt.nets.svs.dstAddr.Prio = raw.nets.svs.dstAddr.Prio.ToString() ?? "";

            // sName – CP1251
            if (raw.nets.svs.sName != null)
                txt.nets.svs.sName = ModbusFunctionsHelper.ConvertByteArrayToString(raw.nets.svs.sName).Replace("\0", "");
            else txt.nets.svs.sName = "";

            txt.nets.svs.nspp = raw.nets.svs.nspp.ToString() ?? "0";
            txt.nets.svs.nasdu = raw.nets.svs.nasdu.ToString() ?? "0";
            txt.nets.svs.nsig = raw.nets.svs.nsig.ToString() ?? "0";
            txt.nets.svs.nasdu_IU = raw.nets.svs.nasdu_IU.ToString() ?? "0";
            txt.nets.svs.gmidentity_ena = raw.nets.svs.gmidentity_ena.ToString() ?? "0";
            txt.nets.svs.smprate_ena = raw.nets.svs.smprate_ena.ToString() ?? "0";

            txt.nets.ips.ipMask = raw.nets.ips.ipMask ?? new byte[0];
            txt.nets.ips.GateWay = raw.nets.ips.GateWay ?? new byte[0];
            txt.nets.ips.ipAddr = raw.nets.ips.ipAddr ?? new byte[0];
            txt.nets.ips.opts = raw.nets.ips.opts ?? new byte[0];

            // name – CP1251
            if (raw.nets.ips.name != null)
                txt.nets.ips.name = ModbusFunctionsHelper.ConvertByteArrayToString(raw.nets.ips.name).Replace("\0", "");
            else txt.nets.ips.name = "";

            // time
            txt.time.ptpval.ns = raw.time.ptpval.ns.ToString() ?? "";
            txt.time.ptpval.slo = raw.time.ptpval.slo.ToString() ?? "";
            txt.time.ptsecHi = raw.time.ptsecHi.ToString() ?? "";
            txt.time.opts = raw.time.opts ?? new byte[0];

            // meas
            txt.meas.primct.Inom1 = raw.meas.primct.Inom1.ToString() ?? "";
            txt.meas.primct.Inom2 = raw.meas.primct.Inom2.ToString() ?? "";
            if (raw.meas.primct.label != null)
                txt.meas.primct.label = ModbusFunctionsHelper.ConvertByteArrayToString(raw.meas.primct.label).Replace("\0", "");
            else txt.meas.primct.label = "";
            txt.meas.Rsh = raw.meas.Rsh.ToString() ?? "";
            txt.meas.sscle = raw.meas.sscle.ToString() ?? "";
            txt.meas.Nturns = raw.meas.Nturns.ToString() ?? "";
            txt.meas.adcrng = raw.meas.adcrng.ToString() ?? "";
            txt.meas.aStart = raw.meas.aStart.ToString() ?? "";
            txt.meas.aStop = raw.meas.aStop.ToString() ?? "";
            txt.meas.adc_osf = raw.meas.adc_osf.ToString();

            // verif
            txt.meas.verif.next_verif = raw.meas.verif.next_verif.ToString();
            txt.meas.verif.verif_st = raw.meas.verif.verif_st.ToString();

            // qb
            txt.meas.qb.method_oor = raw.meas.qb.method_oor.ToString();
            txt.meas.qb.out_of_range = FormatFloatNoExponent(raw.meas.qb.out_of_range);
            txt.meas.qb.delta_od = FormatFloatNoExponent(raw.meas.qb.delta_od);
            txt.meas.qb.repl_smps_od = raw.meas.qb.repl_smps_od.ToString();
            txt.meas.qb.min_lev_od = FormatFloatNoExponent(raw.meas.qb.min_lev_od);
            txt.meas.qb.delta_osc = FormatFloatNoExponent(raw.meas.qb.delta_osc);
            txt.meas.qb.repl_smps_osc = raw.meas.qb.repl_smps_osc.ToString();
            txt.meas.qb.min_lev_osc = FormatFloatNoExponent(raw.meas.qb.min_lev_osc);
            txt.meas.qb.imbal_lev = FormatFloatNoExponent(raw.meas.qb.imbal_lev);
            txt.meas.qb.qb_mod = raw.meas.qb.qb_mod.ToString();
            txt.meas.qb.qb_msk = raw.meas.qb.qb_msk?.Select(v => v.ToString()).ToArray() ?? new string[0];

            txt.nets.gss.cfgRev = raw.nets.gss.cfgRev.ToString();
            txt.nets.gss.dstAddr.dstMAC = raw.nets.gss.dstAddr.dstMAC ?? new byte[0];
            txt.nets.gss.dstAddr.appID = raw.nets.gss.dstAddr.appID.ToString();
            txt.nets.gss.dstAddr.VID = raw.nets.gss.dstAddr.VID.ToString();
            txt.nets.gss.dstAddr.Prio = raw.nets.gss.dstAddr.Prio.ToString();
            txt.nets.gss.sml = raw.nets.gss.sml.ToString();

            // syns
            txt.syns.sevOfs = raw.syns.sevOfs.ToString() ?? "";
            txt.syns.ptps.acv = raw.syns.ptps.acv.ToString() ?? "";
            txt.syns.ptps.ofthi = raw.syns.ptps.ofthi.ToString() ?? "";
            txt.syns.ptps.oftlo = raw.syns.ptps.oftlo.ToString() ?? "";
            txt.syns.ptps.crst = raw.syns.ptps.crst.ToString() ?? "";
            txt.syns.ptps.aclt = raw.syns.ptps.aclt.ToString() ?? "";
            txt.syns.ptps.mmadr = raw.syns.ptps.mmadr ?? new byte[0];
            txt.syns.ptps.rqadr = raw.syns.ptps.rqadr ?? new byte[0];
            txt.syns.ptps.id.portNum = raw.syns.ptps.id.portNum.ToString() ?? "";
            // clkId – CP1251
            if (raw.syns.ptps.id.clkId != null)
                txt.syns.ptps.id.clkId = ModbusFunctionsHelper.ConvertByteArrayToString(raw.syns.ptps.id.clkId).Replace("\0", "");
            else txt.syns.ptps.id.clkId = "";
            txt.syns.ptps.rqpd = raw.syns.ptps.rqpd.ToString() ?? "";
            txt.syns.ptps.mcai = raw.syns.ptps.mcai.ToString() ?? "";
            txt.syns.lmscm = raw.syns.lmscm.ToString() ?? "";
            txt.syns.umscm = raw.syns.umscm.ToString() ?? "";

            // swrcs
            txt.swrcs.swrEna = raw.swrcs.swrEna.ToString() ?? "";
            txt.swrcs.swnf.Inn = raw.swrcs.swnf.Inn.ToString() ?? "";
            txt.swrcs.swnf.Imax = raw.swrcs.swnf.Imax.ToString() ?? "";
            // label – CP1251
            if (raw.swrcs.swnf.label != null)
                txt.swrcs.swnf.label = ModbusFunctionsHelper.ConvertByteArrayToString(raw.swrcs.swnf.label).Replace("\0", "");
            else txt.swrcs.swnf.label = "";
            // model – CP1251
            if (raw.swrcs.swnf.model != null)
                txt.swrcs.swnf.model = ModbusFunctionsHelper.ConvertByteArrayToString(raw.swrcs.swnf.model).Replace("\0", "");
            else txt.swrcs.swnf.model = "";
            txt.swrcs.algo.Iotc = raw.swrcs.algo.Iotc.ToString() ?? "";
            txt.swrcs.algo.Nn = raw.swrcs.algo.Nn.ToString() ?? "";
            txt.swrcs.algo.C1 = FormatFloatNoExponent(raw.swrcs.algo.C1);
            txt.swrcs.algo.C2 = FormatFloatNoExponent(raw.swrcs.algo.C2);
            txt.swrcs.algo.C3 = FormatFloatNoExponent(raw.swrcs.algo.C3);
            txt.swrcs.algo.C4 = FormatFloatNoExponent(raw.swrcs.algo.C4);
            txt.swrcs.jrnvol = raw.swrcs.jrnvol.ToString() ?? "";
            txt.swrcs.contacts.invmsk = raw.swrcs.contacts.invmsk.ToString() ?? "";
            txt.swrcs.contacts.maxdly.offd = raw.swrcs.contacts.maxdly.offd.ToString() ?? "";
            txt.swrcs.contacts.maxdly.ond = raw.swrcs.contacts.maxdly.ond.ToString() ?? "";
            txt.swrcs.contacts.ajtr.offd = raw.swrcs.contacts.ajtr.offd.ToString() ?? "";
            txt.swrcs.contacts.ajtr.ond = raw.swrcs.contacts.ajtr.ond.ToString() ?? "";

            if (raw.swrcs.contacts.cdly == null) raw.swrcs.contacts.cdly = new SCDLY[0];
            txt.swrcs.contacts.cdly = new SCDLY_cdly_TextFormat[raw.swrcs.contacts.cdly.Length];
            for (int i = 0; i < raw.swrcs.contacts.cdly.Length; i++)
            {
                txt.swrcs.contacts.cdly[i].offd = raw.swrcs.contacts.cdly[i].offd;
                txt.swrcs.contacts.cdly[i].ond = raw.swrcs.contacts.cdly[i].ond;
            }

            // accor
            if (raw.accor.AdcOfs == null) raw.accor.AdcOfs = new float[0];
            txt.accor.AdcOfs = new string[raw.accor.AdcOfs.Length];
            for (int i = 0; i < raw.accor.AdcOfs.Length; i++)
                txt.accor.AdcOfs[i] = raw.accor.AdcOfs[i].ToString("R") ?? "0";
            txt.accor.actTbl = raw.accor.actTbl ?? new float[0][];
            txt.accor.phcTbl = raw.accor.phcTbl ?? new uint[0][];

            return txt;
        }

        public void set_GeneralSettings_Text(GeneralSettings_TextFormat value) { }

        public GeneralSettings_TextFormat GeneralSettings_Text_Default => new GeneralSettings_TextFormat();
        public ModBusProfile.esp esp_Default => new ModBusProfile.esp();
        public ModBusProfile.ssp ssp_Default => new ModBusProfile.ssp();
        public List<ModBusProfile.journal_record> journal_record_Default => new List<ModBusProfile.journal_record>();

        public void WriteSettings(GeneralSettings_TextFormat txt, bool AdminRights, Tuple<TcpClient, IModbusMaster> connection)
        {
            if (!AdminRights) return;
            if (connection?.Item1?.Connected != true) return;

            var master = connection.Item2;
            WasRebootCommandSent = false;

            // 1. Запись всех групп настроек в Modbus-буфер устройства
            try
            {
                ModBusProfileHelper.cmns_Write(master, txt, AdminRights);
                ModBusProfileHelper.meas_Write(master, txt, AdminRights);
                ModBusProfileHelper.syns_Write(master, txt, AdminRights);
                ModBusProfileHelper.time_Write(master, txt, AdminRights);
                ModBusProfileHelper.nets_Write(master, txt, AdminRights);
                ModBusProfileHelper.swrcs_Write(master, txt, AdminRights);
                ModBusProfileHelper.accor_Write(master, txt, AdminRights);
            }
            catch (Exception ex)
            {
                if (!IsConnectionError(ex))
                    throw;
                // Если соединение уже разорвано, считаем, что перезагрузка запущена
                WasRebootCommandSent = true;
                return;
            }

            try
            {
                ModBusCommands commands = new ModBusCommands();
                // Отправляем команду перезагрузки. Не ждём ответа, так как устройство закроет соединение.
                commands.upload_settings(master, 0x0100);
                WasRebootCommandSent = true;
            }
            catch (Exception ex)
            {
                if (IsConnectionError(ex))
                    WasRebootCommandSent = true;
                else
                    throw;
            }
        }

        private bool IsConnectionError(Exception ex)
        {
            string msg = ex.Message.ToLower();
            return msg.Contains("socket") ||
                   msg.Contains("connection") ||
                   msg.Contains("disconnected") ||
                   msg.Contains("transport") ||
                   (ex.InnerException != null && IsConnectionError(ex.InnerException));
        }
    }
}