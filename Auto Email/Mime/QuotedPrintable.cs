using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Mime
{
    public class QuotedPrintable
    {
        public static string DecodeEncodedWord(string toDecode, Encoding encoding)
        {
            if (toDecode == null)
                throw new ArgumentNullException("toDecode");

            if (encoding == null)
                throw new ArgumentNullException("encoding");

            // Decode the QuotedPrintable string and return it
            return encoding.GetString(Rfc2047QuotedPrintableDecode(toDecode, true));
        }


        public static byte[] DecodeContentTransferEncoding(string toDecode)
        {
            if (toDecode == null)
                throw new ArgumentNullException("toDecode");

            // Decode the QuotedPrintable string and return it
            return Rfc2047QuotedPrintableDecode(toDecode, false);
        }


        private static byte[] Rfc2047QuotedPrintableDecode(string toDecode, bool encodedWordVariant)
        {
            if (toDecode == null)
                throw new ArgumentNullException("toDecode");

            // Create a byte array builder which is roughly equivalent to a StringBuilder
            using (MemoryStream byteArrayBuilder = new MemoryStream())
            {
                // Remove illegal control characters
                toDecode = RemoveIllegalControlCharacters(toDecode);

                // Run through the whole string that needs to be decoded
                for (int i = 0; i < toDecode.Length; i++)
                {
                    char currentChar = toDecode[i];
                    if (currentChar == '=')
                    {
                        // Check that there is at least two characters behind the equal sign
                        if (toDecode.Length - i < 3)
                        {
                            // We are at the end of the toDecode string, but something is missing. Handle it the way RFC 2045 states
                            WriteAllBytesToStream(byteArrayBuilder, DecodeEqualSignNotLongEnough(toDecode.Substring(i)));

                            // Since it was the last part, we should stop parsing anymore
                            break;
                        }

                        // Decode the Quoted-Printable part
                        string quotedPrintablePart = toDecode.Substring(i, 3);
                        WriteAllBytesToStream(byteArrayBuilder, DecodeEqualSign(quotedPrintablePart));

                        // We now consumed two extra characters. Go forward two extra characters
                        i += 2;
                    }
                    else
                    {
                        // This character is not quoted printable hex encoded.

                        // Could it be the _ character, which represents space
                        // and are we using the encoded word variant of QuotedPrintable
                        if (currentChar == '_' && encodedWordVariant)
                        {
                            // The RFC specifies that the "_" always represents hexadecimal 20 even if the
                            // SPACE character occupies a different code position in the character set in use.
                            byteArrayBuilder.WriteByte(0x20);
                        }
                        else
                        {
                            // This is not encoded at all. This is a literal which should just be included into the output.
                            byteArrayBuilder.WriteByte((byte)currentChar);
                        }
                    }
                }

                return byteArrayBuilder.ToArray();
            }
        }

        private static void WriteAllBytesToStream(Stream stream, byte[] toWrite)
        {
            stream.Write(toWrite, 0, toWrite.Length);
        }


        private static string RemoveIllegalControlCharacters(string input)
        {
            if (input == null)
                throw new ArgumentNullException("input");

            // First we remove any \r or \n which is not part of a \r\n pair
            input = RemoveCarriageReturnAndNewLinewIfNotInPair(input);

            return Regex.Replace(input, "[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "");
        }


        private static string RemoveCarriageReturnAndNewLinewIfNotInPair(string input)
        {
            if (input == null)
                throw new ArgumentNullException("input");

            // Use this for building up the new string. This is used for performance instead
            // of altering the input string each time a illegal token is found
            StringBuilder newString = new StringBuilder(input.Length);

            for (int i = 0; i < input.Length; i++)
            {

                if (input[i] == '\r' && (i + 1 >= input.Length || input[i + 1] != '\n'))
                {

                }
                else if (input[i] == '\n' && (i - 1 < 0 || input[i - 1] != '\r'))
                {
                    // Illegal token \n found. Do not add it to the new string
                }
                else
                {
                    // No illegal tokens found. Simply insert the character we are at
                    // in our new string
                    newString.Append(input[i]);
                }
            }

            return newString.ToString();
        }


        private static byte[] DecodeEqualSignNotLongEnough(string decode)
        {
            if (decode == null)
                throw new ArgumentNullException("decode");

            // We can only decode wrong length equal signs
            if (decode.Length >= 3)
                throw new ArgumentException("decode must have length lower than 3", "decode");

            if (decode.Length <= 0)
                throw new ArgumentException("decode must have length lower at least 1", "decode");

            // First char must be =
            if (decode[0] != '=')
                throw new ArgumentException("First part of decode must be an equal sign", "decode");

            // We will now believe that the string sent to us, was actually not encoded
            // Therefore it must be in US-ASCII and we will return the bytes it corrosponds to
            return Encoding.ASCII.GetBytes(decode);
        }


        private static byte[] DecodeEqualSign(string decode)
        {
            if (decode == null)
                throw new ArgumentNullException("decode");

            // We can only decode the string if it has length 3 - other calls to this function is invalid
            if (decode.Length != 3)
                throw new ArgumentException("decode must have length 3", "decode");

            // First char must be =
            if (decode[0] != '=')
                throw new ArgumentException("decode must start with an equal sign", "decode");


            if (decode.Contains("\r\n"))
            {

                return new byte[0];
            }


            try
            {

                string numberString = decode.Substring(1);

                byte[] oneByte = new[] { Convert.ToByte(numberString, 16) };


                return oneByte;
            }
            catch (FormatException)
            {

                return Encoding.ASCII.GetBytes(decode);
            }
        }
    }
}
