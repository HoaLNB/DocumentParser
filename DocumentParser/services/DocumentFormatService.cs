using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentParser.common;
using DocumentParser.models;
//Have to use this way because in OpenXML.Drawing, 
//there are some class that has the same name as the normal class
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;


namespace DocumentParser.services
{
    /// <summary>
    /// Word Processing utils tool using Open Document SDK
    /// </summary>
    public class DocumentFormatService
    {
        /// <summary>
        /// Get question list from docx file.
        /// </summary>
        /// <param name="docxFilePath"></param>
        /// <param name="imgFolderPath"></param>
        /// <returns></returns>
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

        public void createNewDocxFile(string filePath)
        {
            WordprocessingDocument wordProc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document, true);
            MainDocumentPart mainPart = wordProc.AddMainDocumentPart();
            mainPart.Document = new Document();
            mainPart.Document.AppendChild(new Body());
            wordProc.Close();
        }

        /// <summary>
        /// Export the data to docx file. 
        /// </summary>
        /// <param name="questionList"></param>
        public ResponseResult writeQuestionListToDocx(List<Question> questionList, string docxFilePath)
        {
            ResponseResult responseResult = new ResponseResult();
            try
            {
                WordprocessingDocument wordProc = WordprocessingDocument.Create(docxFilePath,
                    WordprocessingDocumentType.Document, true);
                MainDocumentPart mainPart = wordProc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                foreach (var question in questionList)
                {
                    Table table = new Table();
                    table = createTableFromQuestion(question, mainPart);
                    body.AppendChild(table);
                    body.AppendChild(new Paragraph(new Run(new Text("\n"))));
                }
                wordProc.Close();
                responseResult.RepCode = Constants.RC_EXPORT_SUCCESSFUL;
            }
            catch (IOException ioException)
            {
                //TODO return error code and display correspondent message.
                responseResult.RepCode = Constants.RC_EXPORT_UNABLE_TO_WRITE_FILE;
            }
            return responseResult;
        }

        /// <summary>
        /// Create a table from a question.
        /// </summary>
        /// <param name="question"></param>
        /// <returns></returns>
        public Table createTableFromQuestion(Question question, MainDocumentPart mainPart)
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

            //Adjust the width of the table.
            TableWidth tableWidthProp = new TableWidth();
            tableWidthProp.Type = TableWidthUnitValues.Pct; 
            //1 unit of width is 1/50 of 1 percent -> 5000 unit = 100%
            tableWidthProp.Width = Constants.TABLE_WIDTH_PERCENTAGE_UNIT; 
            table.AppendChild<TableWidth>(tableWidthProp);

            List<StringPair> stringPairListToExport = question.getListOfElement();
            //loop. i is rowCounter
            for (int i = 0; i < stringPairListToExport.Count; i++)
            {
                string firstCellString = stringPairListToExport[i].FirstString;
                string secondCellString = stringPairListToExport[i].SecondString;
                //Append first cell
                var tableRow = new TableRow();
                var firstCell = new TableCell(new Paragraph(new Run(new Text(firstCellString))));
                var secondCell = new TableCell();
                tableRow.AppendChild(firstCell);

                string imagePath = StringUtils.extractImageLinkFromContent(secondCellString);
                Paragraph imageParagraph = new Paragraph();
                if (secondCellString.Contains(Constants.PATTERN_IMG))
                {
                    string imageName =
                        StringUtils.extractImageLinkFromContent(secondCellString)
                            .Substring(StringUtils.extractImageLinkFromContent(secondCellString).LastIndexOf(@"\", StringComparison.Ordinal) + 1);
                    secondCellString = StringUtils.removeImgFromContent(secondCellString) + Constants.PATTERN_IMG +
                                       imageName + Constants.PATTERN_IMG_END_BRACKET;
                    secondCell = new TableCell(new Paragraph(new Run(new Text(secondCellString))));
                    //Insert image into table. Put the image into a paragraph.
                    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                    using (FileStream stream = new FileStream(imagePath, FileMode.Open))
                    {
                        imagePart.FeedData(stream);
                    }
                    imageParagraph = getImageParagraph(mainPart.GetIdOfPart(imagePart));
                    secondCell.AppendChild(imageParagraph);
                }
                else
                {
                    secondCell = new TableCell(new Paragraph(new Run(new Text(secondCellString))));
                }
                tableRow.AppendChild(secondCell);
                table.AppendChild(tableRow);
            }
            return table;
        }
        
        //Borrowed & edited from MSDN's example for inserting an image into a docx using Open XML SDK 2.5. Needs revising.
        private static Paragraph getImageParagraph(string relationshipId)
        {
            Paragraph returnParagraph = new Paragraph();
            // Define the reference of the image.
            var element =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent() {Cx = 3990000L, Cy = 3092000L},
                        new DW.EffectExtent()
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DW.DocProperties()
                        {
                            Id = (UInt32Value) 1U,
                            Name = "Picture 1"
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks() {NoChangeAspect = true}),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties()
                                        {
                                            Id = (UInt32Value) 0U,
                                            Name = "New Bitmap Image.jpg"
                                        },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip(
                                            new A.BlipExtensionList(
                                                new A.BlipExtension()
                                                {
                                                    Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                })
                                            )
                                        {
                                            Embed = relationshipId,
                                            CompressionState =
                                                A.BlipCompressionValues.Print
                                        },
                                        new A.Stretch(
                                            new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset() {X = 0L, Y = 0L },
                                            new A.Extents() {Cx = 3990000L, Cy = 3092000L}),
                                        new A.PresetGeometry(
                                            new A.AdjustValueList()
                                            ) {Preset = A.ShapeTypeValues.Rectangle}))
                                ) {Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"})
                        )
                    {
                        DistanceFromTop = (UInt32Value) 0U,
                        DistanceFromBottom = (UInt32Value) 0U,
                        DistanceFromLeft = (UInt32Value) 0U,
                        DistanceFromRight = (UInt32Value) 0U,
                        EditId = "50D07946"
                    });
            // Append the reference to body, the element should be in a Run.

            returnParagraph = new Paragraph(new Run(element));
            return returnParagraph;
        }
    }
}
