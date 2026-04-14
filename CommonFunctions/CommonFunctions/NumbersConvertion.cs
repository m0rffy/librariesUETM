using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonFunctions
{
    public class NumbersConvertion
    {
        public static string GetHexStringFrom(byte[] byteArray)
        {
            return BitConverter.ToString(byteArray); //To convert the whole array
        }

        public static int HexToInt(string HexValue)
        {
            int IntValue = int.Parse(HexValue, System.Globalization.NumberStyles.HexNumber);
            return IntValue;
        }

        public static long HexToLong(string HexValue)
        {
            long FinalLongValue = 0;
            long LongValue = long.Parse(HexValue, System.Globalization.NumberStyles.HexNumber);
            if ((LongValue >> 31) > 0)
            {
                LongValue = 4294967295 - LongValue;
                LongValue = LongValue * (-1);
                FinalLongValue = LongValue;
            }
            else
            {
                FinalLongValue = LongValue;
            }
            return FinalLongValue;
        }

        public static string HexToText(string HexValue)
        {
            string TextValue = string.Concat(Enumerable
              .Range(0, HexValue.Length / 2)
              .Select(i => (char)int.Parse(HexValue.Substring(2 * i, 2), NumberStyles.HexNumber)));
            return TextValue;
        }

        public static byte[] ConvertHexStringToByteArray(string hexString)
        {
            if (hexString.Length % 2 != 0)
            {
                throw new ArgumentException(String.Format("The binary key cannot have an odd number of digits: {0}", hexString));
            }

            byte[] HexAsBytes = new byte[hexString.Length / 2];
            for (int index = 0; index < HexAsBytes.Length; index++)
            {
                string byteValue = hexString.Substring(index * 2, 2);
                HexAsBytes[index] = byte.Parse(byteValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            }

            return HexAsBytes;
        }

        public static string ConvertStringToHexString(string String)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(String);
            string hexString = BitConverter.ToString(bytes);
            hexString = hexString.Replace("-", "");
            return hexString;
        }

        public static byte[] CombineByteArrays(params byte[][] arrays)
        {
            byte[] ret = new byte[arrays.Sum(x => x.Length)];
            int offset = 0;
            foreach (byte[] data in arrays)
            {
                Buffer.BlockCopy(data, 0, ret, offset, data.Length);
                offset += data.Length;
            }
            return ret;
        }

        public static string ConvertDigitStringToHexFormat(string String)
        {
            if (int.TryParse(String, out int number))
            {
                string hexValue = number.ToString("X");
                return hexValue;
            }
            else
            {
                return String;
            }
        }
    }
}
