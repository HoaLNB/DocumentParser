using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentParser.common;
using DocumentParser.models;

namespace DocumentParser.services
{
    /// <summary>
    /// Word Processing utils tool using Open Document SDK
    /// </summary>
    public class DocumentFormatService
    {
        public List<Question> getQuestionListFromDOCX(string docxFilePath, string imgFolderPath)
        {
            WordprocessingDocument wordProc = WordprocessingDocument.Open(docxFilePath, true);
            Document processedDocument = wordProc.MainDocumentPart.Document;
            List<Question> questionList = new List<Question>();
            List<Table> tableList = processedDocument.Body.Elements<Table>().ToList();
            foreach (var table in tableList)
            {
                questionList.Add(getQuestionFromTable(table, imgFolderPath));
            }
            wordProc.Close();
            return questionList;
        }

        /// <summary>
        /// Get question from a single compatible table in docx
        /// </summary>
        /// <param name="inputTable"></param>
        /// <param name="imgFolderPath"></param>
        /// <returns></returns>
        public Question getQuestionFromTable(Table inputTable, [Optional] string imgFolderPath)
        {
            Question returnedQuestion = new Question();
            foreach (TableRow row in inputTable.Descendants<TableRow>())
            {
                List<TableCell> cellList = row.Elements<TableCell>().ToList();
                if (cellList[0].InnerText.Trim().Contains(Constants.PATTERN_QID))
                {
                    string optImgString = StringUtils.extractImageLinkFromContent(cellList[1].InnerText);
                    if (!string.IsNullOrEmpty(optImgString))
                    {
                        if (!string.IsNullOrEmpty(imgFolderPath))
                        {
                            optImgString = optImgString.Trim().Insert(0, imgFolderPath);
                        }
                    }
                    returnedQuestion.Qid =
                        cellList[0].InnerText.Trim().Substring(cellList[0].InnerText.IndexOf("=") + 1);
                    returnedQuestion.QuestionContent.OptionText =
                        StringUtils.removeImgFromContent(cellList[1].InnerText.Trim());
                    returnedQuestion.QuestionContent.ImageLinkText = optImgString;
                }
                else if (cellList[0].InnerText.Trim().Equals(Constants.CELL_ANSWER))
                {
                    returnedQuestion.Answer = cellList[1].InnerText.Trim();
                }
                else if (cellList[0].InnerText.Trim().Equals(Constants.CELL_MARK))
                {
                    returnedQuestion.Mark = double.Parse(cellList[1].InnerText);
                }
                else if (cellList[0].InnerText.Trim().Equals(Constants.CELL_COURSEID))
                {
                    returnedQuestion.CourseId = cellList[1].InnerText.Trim();
                }
                else if (cellList[0].InnerText.Trim().Equals(Constants.CELL_UNIT))
                {
                    returnedQuestion.Unit = cellList[1].InnerText.Trim();
                }
                else if (cellList[0].InnerText.Trim().Equals(Constants.CELL_MIXCHOICES))
                {
                    if (Constants.YES_CHOICE.Equals(cellList[1].InnerText))
                    {
                        returnedQuestion.MixChoices = true;
                    }
                    else if (Constants.NO_CHOICE.Equals(cellList[1].InnerText))
                    {
                        returnedQuestion.MixChoices = false;
                    }
                }
                else
                {
                    if (cellList[1].InnerText.Trim().Length != 0)
                    {
                        string optImgString = StringUtils.extractImageLinkFromContent(cellList[1].InnerText);
                        if (!string.IsNullOrEmpty(optImgString))
                        {
                            if (!string.IsNullOrEmpty(imgFolderPath))
                            {
                                optImgString = optImgString.Trim().Insert(0, imgFolderPath);
                            }
                        }
                        Option option = new Option();
                        option.OptionText = StringUtils.removeImgFromContent(cellList[1].InnerText.Trim());
                        option.ImageLinkText = optImgString;
                        returnedQuestion.OptionList.Add(option);
                    }
                }
            }
            returnedQuestion.fillOptionList();
            return returnedQuestion;
        }

        /// <summary>
        /// Export the data to docx file. 
        /// </summary>
        /// <param name="questionList"></param>
        public void writeQuestionListToDocx(List<Question> questionList, string docxFilePath)
        {
            WordprocessingDocument wordProc = WordprocessingDocument.Create(docxFilePath, WordprocessingDocumentType.Document, true);
            MainDocumentPart mainPart = wordProc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());
            foreach (var question in questionList)
            {
                Table table = new Table();
                table = createTableFromQuestion(question);
                body.AppendChild(table);
                body.AppendChild(new Paragraph(new Run(new Text("\n"))));
            }
            wordProc.Close();
        }

        /// <summary>
        /// Create a table from a question.
        /// </summary>
        /// <param name="question"></param>
        /// <returns></returns>
        public Table createTableFromQuestion(Question question)
        {
            Table table = new Table();
            TableProperties props = new TableProperties(
            new TableBorders(
            new TopBorder
            {
                Val = new EnumValue<BorderValues>(BorderValues.Single),
                Size = 1
            },
            new BottomBorder
            {
                Val = new EnumValue<BorderValues>(BorderValues.Single),
                Size = 1
            },
            new LeftBorder
            {
                Val = new EnumValue<BorderValues>(BorderValues.Single),
                Size = 1
            },
            new RightBorder
            {
                Val = new EnumValue<BorderValues>(BorderValues.Single),
                Size = 1
            },
            new InsideHorizontalBorder()
            {
                Val = new EnumValue<BorderValues>(BorderValues.Single),
                Size = 1
            },
            new InsideVerticalBorder
            {
                Val = new EnumValue<BorderValues>(BorderValues.Single),
                Size = 1
            }));
            table.AppendChild<TableProperties>(props);
            List<StringPair> stringPairListToExport = question.getListOfElement();
            //loop. i is rowCounter
            for (int i = 0; i < stringPairListToExport.Count; i++)
            {
                string firstCellString = stringPairListToExport[i].FirstString;
                string secondCellString = stringPairListToExport[i].SecondString;
                if (secondCellString.Contains(Constants.PATTERN_IMG))
                {
                    string imagePath = StringUtils.extractImageLinkFromContent(secondCellString).Substring(StringUtils.extractImageLinkFromContent(secondCellString).LastIndexOf(@"\")+1);
                    secondCellString = StringUtils.removeImgFromContent(secondCellString) + Constants.PATTERN_IMG + imagePath + Constants.PATTERN_IMG_END_BRACKET;
                }
                var tableRow = new TableRow();
                var firstCell = new TableCell(new Paragraph(new Run(new Text(firstCellString))));
                tableRow.AppendChild(firstCell);
                var secondCell = new TableCell(new Paragraph(new Run(new Text(secondCellString))));
                tableRow.AppendChild(secondCell);
                table.AppendChild(tableRow);
            }
            return table;
        }
    }
}
