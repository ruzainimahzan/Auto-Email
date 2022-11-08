using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using OpenPop.Common.Logging;

namespace Mime
{
    public class Base64
    {
        public static byte[] Decode(string base64Encoded)
        {

            try
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    base64Encoded = base64Encoded.Replace("\r\n", "");
                    base64Encoded = base64Encoded.Replace("\t", "");
                    base64Encoded = base64Encoded.Replace(" ", "");

                    byte[] inputBytes = Encoding.ASCII.GetBytes(base64Encoded);

                    using (FromBase64Transform transform = new FromBase64Transform(FromBase64TransformMode.DoNotIgnoreWhiteSpaces))
                    {
                        byte[] outputBytes = new byte[transform.OutputBlockSize];


                        const int inputBlockSize = 4;
                        int currentOffset = 0;
                        while (inputBytes.Length - currentOffset > inputBlockSize)
                        {
                            transform.TransformBlock(inputBytes, currentOffset, inputBlockSize, outputBytes, 0);
                            currentOffset += inputBlockSize;
                            memoryStream.Write(outputBytes, 0, transform.OutputBlockSize);
                        }


                        outputBytes = transform.TransformFinalBlock(inputBytes, currentOffset, inputBytes.Length - currentOffset);
                        memoryStream.Write(outputBytes, 0, outputBytes.Length);
                    }

                    return memoryStream.ToArray();
                }
            }
            catch (FormatException e)
            {
                DefaultLogger.Log.LogError("Base64: (FormatException) " + e.Message + "\r\nOn string: " + base64Encoded);
                throw;
            }
        }


        public static string Decode(string base64Encoded, Encoding encoding)
        {
            if (base64Encoded == null)
                throw new ArgumentNullException("base64Encoded");

            if (encoding == null)
                throw new ArgumentNullException("encoding");

            return encoding.GetString(Decode(base64Encoded));
        }
    }
}
