using System.Linq;

namespace DocumentParser.common
{
    public static class StringUtils
    {
        /// <summary>
        /// Find next letter alphabetically.
        /// </summary>
        /// <param name="inputLetter"></param>
        /// <returns></returns>
        public static char alphabetIncrement(char inputLetter)
        {
            char nextLetter;
            if (inputLetter == 'z') nextLetter = 'a';
            else if (inputLetter == 'Z') nextLetter = 'A';
            else nextLetter = (char) (inputLetter + 1);
            return nextLetter;
        }

        /// <summary>
        /// Extract image link from option's content (in docx file).
        /// </summary>
        /// <param name="content"></param>
        /// <returns></returns>
        public static string extractImageLinkFromContent(string content)
        {
            string optImgString = string.Empty;
            if (content.Contains(Constants.PATTERN_IMG))
            {
                optImgString = string.Concat(content.Substring(content.IndexOf(Constants.PATTERN_IMG) + 6).TakeWhile((c) => c != ']'));
            }
            return optImgString;
        }

        /// <summary>
        /// Remove image link from option content (to insert into database).
        /// </summary>
        /// <param name="content"></param>
        /// <returns></returns>
        public static string removeImgFromContent(string content)
        {
            string returnString = content;
            string optImgString = string.Empty;
            if (content.Contains(Constants.PATTERN_IMG))
            {
                optImgString = string.Concat(content.Substring(content.IndexOf(Constants.PATTERN_IMG) + 6).TakeWhile((c) => c != ']'));
                returnString = content.Replace(Constants.PATTERN_IMG + optImgString + Constants.PATTERN_IMG_END_BRACKET,string.Empty);
            }
            return returnString;
        }

    }
}