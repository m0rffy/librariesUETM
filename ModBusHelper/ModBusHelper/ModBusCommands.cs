using NModbus;
using NModbus.Message;
using System.Threading;

namespace ModBusHelper
{
    public class ModBusCommands
    {
        private ModBusFunctions ModbusFunctionsHelper = new ModBusFunctions();

        public bool WaitForWriteComplete(IModbusMaster master, int timeoutMs = 10000)
        {
            ushort statusAddr = 12290; // 0x3002
            int elapsed = 0;
            int delay = 200;
            while (elapsed < timeoutMs)
            {
                bool[] coils = master.ReadCoils(0, statusAddr, 8);
                byte status = 0;
                for (int i = 0; i < 8; i++)
                    if (coils[i]) status |= (byte)(1 << i);

                if (status == 0x00) // OK
                    return true;

                Thread.Sleep(delay);
                elapsed += delay;
            }
            return false;
        }

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