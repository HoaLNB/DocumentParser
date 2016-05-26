using System;
using DocumentParser.services;
using Microsoft.Office.Interop.Word;

namespace DocumentParser
{
    /// <summary>
    /// Word Processing utils tool using Office Interop
    /// </summary>
    public class ExperimentalServiceInterop
    {
        public void doSomething()
        {
            // Open a doc file.
            Application application = new Application();
            Document document = application.Documents.Open(@"D:\\DocumentParser\\TargetDoc.docx");

            // Loop through all words in the document.
            int count = document.Words.Count;

            for (int i = 1; i <= count; i++)
            {
                // Write the word.
                string text = document.Words[i].Text;
                Console.WriteLine("Word {0} = {1}", i, text);
            }
            // Close word.
            application.Quit();
        }

        public void doSomething2()
        {
            DocumentFormatService dfs = new DocumentFormatService();
            dfs.createNewDocxFile(@"D:\Experiment.docx");
        }
    }
}
