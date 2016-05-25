namespace DocumentParser.models
{
    /// <summary>
    /// Class for each option including text and image link.
    /// </summary>
    public class Option
    {
        private string optionText;
        private string imageLinkText;

        public Option(string optionText, string imageLinkText)
        {
            this.optionText = optionText;
            this.imageLinkText = imageLinkText;
        }

        public Option()
        {
            this.imageLinkText = string.Empty;
        }

        public string OptionText
        {
            get { return optionText; }
            set { optionText = value; }
        }

        public string ImageLinkText
        {
            get { return imageLinkText; }
            set { imageLinkText = value; }
        }
    }
}