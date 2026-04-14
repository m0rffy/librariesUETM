using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Globalization;
using NModbus;

namespace ModBusHelper
{
    public class ModBusFunctions
    {

        public static bool ValidateStringLength(string text, int maxBytes, out int actualBytes)
        {
            actualBytes = Encoding.UTF8.GetByteCount(text ?? "");
            return actualBytes <= maxBytes - 1; // -1 для нуль-терминатора
        }

        public ushort[] GetReadOnlyRegisters(IModbusMaster Master, ushort StartRegister, ushort NumberOfRegistersToBeRead)
        {
            return Master.ReadInputRegisters(0, StartRegister, NumberOfRegistersToBeRead);
        }

        public ushort[] GetRWRegisters(IModbusMaster Master, ushort StartRegister, ushort NumberOfRegistersToBeRead)
        {
            return Master.ReadHoldingRegisters(0, StartRegister, NumberOfRegistersToBeRead);
        }

        public void WriteSingleRegister(IModbusMaster Master, ushort RegisterAddress, ushort Value)
        {
            Master.WriteSingleRegister(0, RegisterAddress, Value);
        }

        public void WriteSingleRegisterThroughLimitation(IModbusMaster Master, ushort RegisterAddress, ushort Value)
        {
            List<ushort> list = new List<ushort> { Value };
            Master.WriteMultipleRegisters(0, RegisterAddress, list.ToArray());
        }

        public void WriteMuiltipleRegisters(IModbusMaster Master, ushort StartRegister, ushort[] Data)
        {
            Master.WriteMultipleRegisters(0, StartRegister, Data);
        }

        public bool[] ReadCoils(IModbusMaster Master, ushort StartRegister, ushort NumberOfRegistersToBeRead)
        {
            return Master.ReadCoils(0, StartRegister, NumberOfRegistersToBeRead);
        }

        public ushort[] ConvertStringToUshortArray(string text, int byteArrayLength)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(text);
            if (bytes.Length < byteArrayLength)
            {
                Array.Resize(ref bytes, byteArrayLength);
                for (int i = bytes.Length; i < byteArrayLength; i++) bytes[i] = 0;
            }
            else if (bytes.Length > byteArrayLength)
            {
                Array.Resize(ref bytes, byteArrayLength);
            }
            return ConvertByteArrayToUShortArray(bytes, true);
        }

        public string ConvertUshortArrayToString(ushort[] registers)
        {
            byte[] bytes = ConvertUshortArrayToBytes(registers);
            int len = Array.IndexOf(bytes, (byte)0);
            if (len < 0) len = bytes.Length;
            return Encoding.UTF8.GetString(bytes, 0, len);
        }

        public byte[] ConvertUshortArrayToBytes(ushort[] UshortArray)
        {
            List<byte> list = new List<byte>();
            foreach (ushort value in UshortArray)
            {
                byte[] bytes = BitConverter.GetBytes(value);
                Array.Reverse(bytes);
                list.AddRange(bytes);
            }
            return list.ToArray();
        }

        public sbyte[] ConvertUshortArrayToSbyteArray(ushort[] ushortArray)
        {
            sbyte[] array = new sbyte[ushortArray.Length * 2];
            for (int i = 0; i < ushortArray.Length; i++)
            {
                byte[] bytes = BitConverter.GetBytes(ushortArray[i]);
                array[i * 2] = (sbyte)bytes[0];
                array[i * 2 + 1] = (sbyte)bytes[1];
            }
            return array;
        }

        public byte[] ConvertUshortToBytes(ushort Ushort)
        {
            return BitConverter.GetBytes(Ushort);
        }

        public static double GetDouble(ushort b3, ushort b2, ushort b1, ushort b0)
        {
            byte[] value = BitConverter.GetBytes(b0)
                .Concat(BitConverter.GetBytes(b1))
                .Concat(BitConverter.GetBytes(b2))
                .Concat(BitConverter.GetBytes(b3))
                .ToArray();
            return BitConverter.ToDouble(value, 0);
        }

        public float GetSingle(ushort high, ushort low)
        {
            byte[] byteArray = new byte[4];
            if (BitConverter.IsLittleEndian)
            {
                byteArray[0] = (byte)(low & 0xFF);
                byteArray[1] = (byte)(low >> 8);
                byteArray[2] = (byte)(high & 0xFF);
                byteArray[3] = (byte)(high >> 8);
            }
            else
            {
                byteArray[3] = (byte)(low & 0xFF);
                byteArray[2] = (byte)(low >> 8);
                byteArray[1] = (byte)(high & 0xFF);
                byteArray[0] = (byte)(high >> 8);
            }
            return BitConverter.ToSingle(byteArray, 0);
        }

        public uint GetUInt32(ushort highOrderValue, ushort lowOrderValue)
        {
            return BitConverter.ToUInt32(BitConverter.GetBytes(lowOrderValue).Concat(BitConverter.GetBytes(highOrderValue)).ToArray(), 0);
        }

        public int GetInt32(ushort highOrderValue, ushort lowOrderValue)
        {
            return BitConverter.ToInt32(BitConverter.GetBytes(lowOrderValue).Concat(BitConverter.GetBytes(highOrderValue)).ToArray(), 0);
        }

        public byte[] GetAsciiBytes(params byte[] numbers)
        {
            return Encoding.ASCII.GetBytes(numbers.SelectMany(n => n.ToString("X2")).ToArray());
        }

        public byte[] GetAsciiBytes(params ushort[] numbers)
        {
            return Encoding.ASCII.GetBytes(numbers.SelectMany(n => n.ToString("X4")).ToArray());
        }

        public ushort[] NetworkBytesToHostUInt16(byte[] networkBytes)
        {
            ushort[] result = new ushort[networkBytes.Length / 2];
            for (int i = 0; i < result.Length; i++)
                result[i] = (ushort)IPAddress.NetworkToHostOrder(BitConverter.ToInt16(networkBytes, i * 2));
            return result;
        }

        public byte[] HexToBytes(string hex)
        {
            byte[] bytes = new byte[hex.Length / 2];
            for (int i = 0; i < bytes.Length; i++)
                bytes[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
            return bytes;
        }

        public uint ConvertUshortArrayToUint(ushort[] UshortArray)
        {
            uint high = (uint)UshortArray[0] << 16;
            uint low = UshortArray[1];
            return high | low;
        }

        public string HexToText(string HexValue)
        {
            return string.Concat(Enumerable.Range(0, HexValue.Length / 2)
                .Select(i => (char)int.Parse(HexValue.Substring(2 * i, 2), NumberStyles.HexNumber)));
        }

        public ushort[] ConvertUintToUshortArray(uint value)
        {
            ushort[] result = new ushort[2];
            byte[] bytes = BitConverter.GetBytes(value);
            if (BitConverter.IsLittleEndian)
            {
                result[0] = BitConverter.ToUInt16(bytes, 0);
                result[1] = BitConverter.ToUInt16(bytes, 2);
            }
            else
            {
                result[0] = BitConverter.ToUInt16(bytes, 2);
                result[1] = BitConverter.ToUInt16(bytes, 0);
            }
            return result;
        }

        public ushort[] ConvertIntToUshortArray(int value)
        {
            ushort[] result = new ushort[2];
            if (BitConverter.IsLittleEndian)
            {
                result[0] = (ushort)(value & 0xFFFF);
                result[1] = (ushort)((value >> 16) & 0xFFFF);
            }
            else
            {
                result[1] = (ushort)(value & 0xFFFF);
                result[0] = (ushort)((value >> 16) & 0xFFFF);
            }
            return result;
        }

        private long GetSignedIntegerRepresentation(float value)
        {
            long num = BitConverter.DoubleToInt64Bits(value);
            if ((ulong)num > 4294967295uL) num--;
            return num;
        }

        public ushort[] ConvertFloatToUshortArray(float number)
        {
            byte[] bytes = BitConverter.GetBytes(number);
            if (BitConverter.IsLittleEndian)
                return new ushort[] { BitConverter.ToUInt16(bytes, 0), BitConverter.ToUInt16(bytes, 2) };
            else
            {
                Array.Reverse(bytes);
                return new ushort[] { BitConverter.ToUInt16(bytes, 0), BitConverter.ToUInt16(bytes, 2) };
            }
        }

        public ushort ConvertShortToUshort(short Short)
        {
            return BitConverter.ToUInt16(BitConverter.GetBytes(Short), 0);
        }

        public ushort[] ConvertAsciiStringToUshortArray(string AsciiString, int AsciiByteArrayProfileLength)
        {
            List<byte> list = new List<byte>();
            byte[] bytes = Encoding.ASCII.GetBytes(AsciiString);
            list.AddRange(bytes);
            for (int i = 0; i < AsciiByteArrayProfileLength - AsciiString.Length; i++)
                list.Add(0);
            return ConvertByteArrayToUShortArray(list.ToArray(), true);
        }

        public ushort[] ConvertByteArrayToUshortArray(byte[] ByteArray)
        {
            ushort[] array = new ushort[ByteArray.Length / 2];
            int num = 0;
            for (int i = 0; i < array.Length; i += 2)
            {
                ushort num2 = BitConverter.ToUInt16(ByteArray, i);
                array[num] = num2;
                num++;
            }
            return array;
        }

        public ushort[] ConvertHexStringArrayToUshortArray2(string[] StringArray)
        {
            byte[] array = new byte[StringArray.Length];
            for (int i = 0; i < StringArray.Length; i++)
                array[i] = Convert.ToByte(StringArray[i], 16);
            return ConvertByteArrayToUshortArray(array);
        }

        public ushort[] ConvertHexStringArrayToUshortArray(string[] hexStrings, bool LittleEndian, bool NetworkOrder)
        {
            ushort[] array = new ushort[3];
            if (LittleEndian)
            {
                for (int i = 0; i < array.Length; i++)
                    array[i] = Convert.ToUInt16(hexStrings[i * 2] + hexStrings[i * 2 + 1], 16);
            }
            else
            {
                for (int i = 0; i < array.Length; i++)
                    array[i] = Convert.ToUInt16(hexStrings[i * 2 + 1] + hexStrings[i * 2], 16);
            }
            if (NetworkOrder) Array.Reverse(array);
            return array;
        }

        public ushort[] ConvertByteArrayToUShortArray(byte[] macBytes, bool NetworkOrder)
        {
            if (macBytes == null || macBytes.Length == 0)
                return new ushort[0];

            ushort[] array = new ushort[macBytes.Length / 2];
            for (int i = 0; i < array.Length; i++)
            {
                if (NetworkOrder)
                    array[i] = (ushort)((macBytes[i * 2] << 8) | macBytes[i * 2 + 1]);
                else
                    array[i] = (ushort)((macBytes[i * 2 + 1] << 8) | macBytes[i * 2]);
            }
            return array;
        }

        public byte[] ConvertHexStringArrayToByteArray(string[] hexStrings)
        {
            byte[] bytes = new byte[hexStrings.Length];
            for (int i = 0; i < hexStrings.Length; i++)
                bytes[i] = byte.Parse(hexStrings[i], NumberStyles.HexNumber, CultureInfo.CurrentCulture);
            return bytes;
        }

        public ushort[] ConvertTextFloatToUshortArray(string StringWithFloat)
        {
            float number = float.Parse(StringWithFloat, CultureInfo.CurrentCulture.NumberFormat);
            return ConvertFloatToUshortArray(number);
        }

        public ushort[] ConvertTextUintArrayToUshortArray(string[] StringUintArray)
        {
            uint[] array = new uint[StringUintArray.Length];
            for (int i = 0; i < array.Length; i++)
                array[i] = uint.Parse(StringUintArray[i]);
            ushort[] array2 = new ushort[array.Length * 2];
            int num = 0;
            for (int j = 0; j < array.Length; j++)
            {
                ushort[] temp = ConvertUintToUshortArray(array[j]);
                array2[num] = temp[0];
                array2[num + 1] = temp[1];
                num += 2;
            }
            return array2;
        }

        public byte ConvertBoolArrayToByte(bool[] source)
        {
            byte result = 0;
            int index = 8 - source.Length;
            foreach (bool b in source)
            {
                if (b) result |= (byte)(1 << (7 - index));
                index++;
            }
            return result;
        }

        public string ConvertByteArrayToString(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0) return "";
            int len = Array.IndexOf(bytes, (byte)0);
            if (len < 0) len = bytes.Length;
            return Encoding.UTF8.GetString(bytes, 0, len);
        }

    }
}