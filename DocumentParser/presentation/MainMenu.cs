using System;
using System.Collections.Generic;
using System.Windows.Forms;
using DocumentParser.common;
using DocumentParser.services;

namespace DocumentParser.presentation
{
    public partial class MainMenu : Form
    {
        public MainMenu()
        {
            StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();
        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.ShowDialog();
            string selectedFolder = folderBrowserDialog.SelectedPath;
            if (!selectedFolder.EndsWith(@"\")) selectedFolder += @"\";
            txtImgFolderPath.Text = selectedFolder;
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            FileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = Constants.WORD_FILE_DIALOG_FILTER;
            openFileDialog.FilterIndex = 1;

            openFileDialog.ShowDialog();
            string selectedFileName = openFileDialog.FileName;
            txtFilePath.Text = selectedFileName;
        }

        private void btImport_Click(object sender, EventArgs e)
        {
            if (!txtImgFolderPath.Text.Trim().EndsWith(@"\")) txtImgFolderPath.Text += @"\";
            string docxFilePath = txtFilePath.Text.Trim();
            string imgFolderPath = txtImgFolderPath.Text.Trim();
            if (QuestionService.getQuestionListFromFile(docxFilePath, imgFolderPath).Count==0)
            {
                MessageBox.Show(Constants.MSG_MB_NO_QUESTION);
                return;
            }
            ResponseResult responseResult = QuestionService.insertFromQuestionList(QuestionService.getQuestionListFromFile(docxFilePath, imgFolderPath));
            if (Constants.RC_SQL_INSERT_ALL_SUCCESSFUL.Equals(responseResult.RepCode))
            {
                MessageBox.Show(Constants.MSG_MB_INSERT_ALL_SUCCESSFUL);
            }
            else
            {
                if (Constants.RC_SQL_INSERT_PARTIAL_PRIMARY_KEY.Equals(responseResult.RepCode))
                {
                    string messageToShow = Constants.MSG_MB_INSERT_PARTIAL_FAIL;
                    foreach (var qid in (List<string>) responseResult.RepData)
                    {
                        messageToShow += qid + " ";
                    }
                    messageToShow += Constants.MSG_MB_INSERT_PARTIAL_FAIL_AFTER;
                    MessageBox.Show(messageToShow);
                }
                else if (Constants.RC_SQL_INSERT_ALL_FAIL_PRIMARY_KEY.Equals(responseResult.RepCode))
                {
                    MessageBox.Show(Constants.MSG_MB_INSERT_ALL_FAIL);
                }
            }
        }

        private void btCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btExport_Click(object sender, EventArgs e)
        {
            FileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = Constants.WORD_FILE_DIALOG_FILTER;
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.ShowDialog();

            string docxFilePath = saveFileDialog.FileName;
            QuestionService.exportQuestionDatabaseToDocx(docxFilePath);
        }
    }
}
