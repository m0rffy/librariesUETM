using NModbus.Device;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using CommonFunctions;
using System.Net;
using System.Net.Http;
using System.Runtime.Remoting.Messaging;
using static ModBusHelper.ModBusExporterLinker;
using System.Globalization;
using System.Xml.Linq;
using NModbus.Extensions.Enron;
using System.Diagnostics.Contracts;
using static ModBusHelper.ModBusProfile;
using System.Security.Cryptography;
using NModbus.Message;
using NModbus;

namespace ModBusHelper
{
    public class ModBusCommands
    {
        private ModBusFunctions ModbusFunctionsHelper = new ModBusFunctions();

        public WriteSingleRegisterRequestResponse upload_firmware_default(IModbusMaster Master, ushort bit_mask)
        {
            ushort startAddress = 12288;
            return Master.ExecuteCustomMessage<WriteSingleRegisterRequestResponse>(new WriteSingleRegisterRequestResponse(0, startAddress, bit_mask));
        }

        public WriteSingleRegisterRequestResponse upload_settings(IModbusMaster Master, ushort reset_flag)
        {
            ushort startAddress = 12289;
            return Master.ExecuteCustomMessage<WriteSingleRegisterRequestResponse>(new WriteSingleRegisterRequestResponse(0, startAddress, reset_flag));
        }

        public bool[] get_last_settings_record(IModbusMaster Master)
        {
            ushort startRegister = 12290;
            ushort numberOfRegistersToBeRead = 8;
            return ModbusFunctionsHelper.ReadCoils(Master, startRegister, numberOfRegistersToBeRead);
        }

        public WriteSingleRegisterRequestResponse reset_ect(IModbusMaster Master)
        {
            ushort startAddress = 12291;
            return Master.ExecuteCustomMessage<WriteSingleRegisterRequestResponse>(new WriteSingleRegisterRequestResponse(0, startAddress, 255));
        }

        public WriteSingleRegisterRequestResponse nulify_swrc(IModbusMaster Master)
        {
            ushort startAddress = 12292;
            return Master.ExecuteCustomMessage<WriteSingleRegisterRequestResponse>(new WriteSingleRegisterRequestResponse(0, startAddress, 255));
        }

        public WriteSingleRegisterRequestResponse clear_swrc_journal(IModbusMaster Master, bool clear_journal)
        {
            ushort startAddress = 12293;
            if (clear_journal)
                return Master.ExecuteCustomMessage<WriteSingleRegisterRequestResponse>(new WriteSingleRegisterRequestResponse(0, startAddress, 1));
            else
                return Master.ExecuteCustomMessage<WriteSingleRegisterRequestResponse>(new WriteSingleRegisterRequestResponse(0, startAddress, 0));
        }
    }
}