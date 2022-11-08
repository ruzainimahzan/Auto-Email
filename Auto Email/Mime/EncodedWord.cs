using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Mime
{
	public class EncodedWord
	{
        public static string Decode(string encodedWords)
        {
            if (encodedWords == null)
                throw new ArgumentNullException("encodedWords");

            const string encodedWordRegex = @"\=\?(?<Charset>\S+?)\?(?<Encoding>\w)\?(?<Content>.+?)\?\=";

            const string replaceRegex = @"(?<first>" + encodedWordRegex + @")\s+(?<second>" + encodedWordRegex + ")";


            encodedWords = Regex.Replace(encodedWords, replaceRegex, "${first}${second}");
            encodedWords = Regex.Replace(encodedWords, replaceRegex, "${first}${second}");

            string decodedWords = encodedWords;

            MatchCollection matches = Regex.Matches(encodedWords, encodedWordRegex);
            foreach (Match match in matches)
            {

                if (!match.Success) continue;

                string fullMatchValue = match.Value;

                string encodedText = match.Groups["Content"].Value;
                string encoding = match.Groups["Encoding"].Value;
                string charset = match.Groups["Charset"].Value;


                Encoding charsetEncoding = EncodingFinder.FindEncoding(charset);

                string decodedText;

                switch (encoding.ToUpperInvariant())
                {

                    case "B":
                        decodedText = Base64.Decode(encodedText, charsetEncoding);
                        break;

                    case "Q":
                        decodedText = QuotedPrintable.DecodeEncodedWord(encodedText, charsetEncoding);
                        break;

                    default:
                        throw new ArgumentException("The encoding " + encoding + " was not recognized");
                }


                decodedWords = decodedWords.Replace(fullMatchValue, decodedText);
            }

            return decodedWords;
        }

	}
}
