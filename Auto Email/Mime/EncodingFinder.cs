using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mime
{
    public class EncodingFinder
    {
        public delegate Encoding FallbackDecoderDelegate(string characterSet);

        public static FallbackDecoderDelegate FallbackDecoder { private get; set; }

        private static Dictionary<string, Encoding> EncodingMap { get; set; }

        static EncodingFinder()
        {
            Reset();
        }

        internal static void Reset()
        {
            EncodingMap = new Dictionary<string, Encoding>();
            FallbackDecoder = null;


            AddMapping("utf8", Encoding.UTF8);
            AddMapping("binary", Encoding.ASCII);
        }

        internal static Encoding FindEncoding(string characterSet)
        {
            if (characterSet == null)
                throw new ArgumentNullException("characterSet");

            string charSetUpper = characterSet.ToUpperInvariant();

            if (EncodingMap.ContainsKey(charSetUpper))
                return EncodingMap[charSetUpper];

            try
            {
                if (charSetUpper.Contains("WINDOWS") || charSetUpper.Contains("CP"))
                {
                    charSetUpper = charSetUpper.Replace("CP", "");
                    charSetUpper = charSetUpper.Replace("WINDOWS", "");
                    charSetUpper = charSetUpper.Replace("-", "");

                    int codepageNumber = int.Parse(charSetUpper, CultureInfo.InvariantCulture);

                    return Encoding.GetEncoding(codepageNumber);
                }

                if (charSetUpper.Length > 3 && charSetUpper.StartsWith("ISO") && charSetUpper[3] >= '0' && charSetUpper[3] <= '9')
                {
                    return Encoding.GetEncoding("iso-" + characterSet.Substring(3));
                }

                return Encoding.GetEncoding(characterSet);
            }
            catch (ArgumentException)
            {
                // The encoding could not be found generally. 
                // Try to use the FallbackDecoder if it is defined.

                // Check if it is defined
                if (FallbackDecoder == null)
                    throw; // It was not defined - throw catched exception

                // Use the FallbackDecoder
                Encoding fallbackDecoderResult = FallbackDecoder(characterSet);

                // Check if the FallbackDecoder had a solution
                if (fallbackDecoderResult != null)
                    return fallbackDecoderResult;

                // If no solution was found, throw catched exception
                throw;
            }
        }

        public static void AddMapping(string characterSet, Encoding encoding)
        {
            if (characterSet == null)
                throw new ArgumentNullException("characterSet");

            if (encoding == null)
                throw new ArgumentNullException("encoding");

            // Add the mapping using uppercase
            EncodingMap.Add(characterSet.ToUpperInvariant(), encoding);
        }
    }
}
