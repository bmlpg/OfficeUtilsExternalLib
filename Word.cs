using System.Collections;
using System.Text;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XWPF.Extractor;
using NPOI.XWPF.UserModel;
using SixLabors.ImageSharp;

namespace OfficeUtilsExternalLib
{
    internal class Word
    {

        private enum WordOutputType : int
        {
            Text = 1,
            LegacyTable = 2,
            Picture = 3,
            Table = 4
        }

        public static byte[] GenerateWordFile(WordStructures.WordFile wordFile)
        {
            XWPFDocument document = null;

            if (wordFile.Template.Length > 0)
            {
                MemoryStream templateStream = new MemoryStream(wordFile.Template);
                document = new XWPFDocument(templateStream);
            }
            else
            {
                throw new Exception("Word Template Undefined");
            }

            List<XWPFParagraph> ParagraphsToRemove = new List<XWPFParagraph>();
            List<XWPFTable> TablesToRemove = new List<XWPFTable>();

            //Document
            IEnumerator documentParagraphsIterator = document.GetParagraphsEnumerator();

            ProcessParagraphs(documentParagraphsIterator, wordFile, ParagraphsToRemove);

            if (ParagraphsToRemove.Count != 0)
            {
                foreach (XWPFParagraph p in ParagraphsToRemove)
                {
                    document.RemoveBodyElement(document.GetPosOfParagraph(p));
                }
                ParagraphsToRemove.Clear();
            }

            IEnumerator documentTablesIterator = document.GetTablesEnumerator();
            ProcessTables(documentTablesIterator, wordFile, TablesToRemove);

            if (TablesToRemove.Count != 0)
            {
                foreach (XWPFTable t in TablesToRemove)
                {
                    document.RemoveBodyElement(document.GetPosOfTable(t));
                }
                TablesToRemove.Clear();
            }

            //Header
            IEnumerator headerEnumerator = document.HeaderList.GetEnumerator();

            while (headerEnumerator.MoveNext())
            {
                XWPFHeader header = (XWPFHeader)headerEnumerator.Current;

                IEnumerator headerParagraphsEnumerator = header.GetListParagraph().GetEnumerator();
                ProcessParagraphs(headerParagraphsEnumerator, wordFile, ParagraphsToRemove);

                IEnumerator headerTablesEnumerator = header.Tables.GetEnumerator();
                ProcessTables(headerTablesEnumerator, wordFile, TablesToRemove);
            }


            //Footer
            IEnumerator footerEnumerator = document.FooterList.GetEnumerator();

            while (footerEnumerator.MoveNext())
            {
                XWPFFooter footer = (XWPFFooter)footerEnumerator.Current;

                IEnumerator footerParagraphsEnumerator = footer.GetListParagraph().GetEnumerator();
                ProcessParagraphs(footerParagraphsEnumerator, wordFile, ParagraphsToRemove);

                IEnumerator footerTablesEnumerator = footer.Tables.GetEnumerator();
                ProcessTables(footerTablesEnumerator, wordFile, TablesToRemove);
            }

            MemoryStream outputStream = new MemoryStream();
            document.Write(outputStream);
            byte[] wordBinary = outputStream.ToArray();

            return wordBinary;

        } // MssGenerateWordFile

        private static void ProcessParagraphs(IEnumerator paragraphsEnumerator, WordStructures.WordFile wordFile, List<XWPFParagraph> ParagraphsToRemove)
        {
            while (paragraphsEnumerator.MoveNext())
            {
                XWPFParagraph paragraph = (XWPFParagraph)paragraphsEnumerator.Current;
                ProcessParagraph(paragraph, wordFile, ParagraphsToRemove);
            }
        }
        private static void ProcessParagraph(XWPFParagraph paragraph, WordStructures.WordFile wordFile, List<XWPFParagraph> ParagraphsToRemove)
        {
            for (int i = 0; i < wordFile.WordOutputs.Count; i++)
            {
                WordStructures.WordOutput wordOutput = wordFile.WordOutputs[i];

                if (wordOutput.OutputType == (int)WordOutputType.Text || wordOutput.OutputType == (int)WordOutputType.Picture)
                {
                    int matchIndex = HasMatch(paragraph, wordOutput.Placeholder);
                    if (matchIndex >= 0)
                    {
                        if (wordOutput.DeletePlaceholder)
                        {
                            //If the placeholder is the only text in all paragraph then paragraph should be removed, if there is additional text - placeholder should be replaced by empty string.
                            if (paragraph.ParagraphText == wordOutput.Placeholder)
                            {
                                ParagraphsToRemove.Add(paragraph);
                            }
                            else
                            {
                                ReplaceTextWithText(paragraph, wordOutput.Placeholder, matchIndex, new WordStructures.WordText());
                            }
                        }
                        else
                        {

                            if (wordOutput.OutputType == (int)WordOutputType.Text)
                            {
                                wordOutput.WordText.Text = Utils.VerifyTextForXML(wordOutput.WordText.Text, "Placeholder:'" + wordOutput.Placeholder + "'", wordFile.Options.AutoRemoveInvalidXMLChars);
                                ReplaceTextWithText(paragraph, wordOutput.Placeholder, matchIndex, wordOutput.WordText);
                            }
                            else if (wordOutput.OutputType == (int)WordOutputType.Picture)
                            {
                                ReplaceTextWithPicture(paragraph, wordOutput.Placeholder, matchIndex, wordOutput.WordPicture);
                            }
                        }
                    }
                    else if (wordOutput.OutputType == (int)WordOutputType.Text)
                    {
                        //replace text inside text box   
                        IEnumerator runsEnumerator = paragraph.Runs.GetEnumerator();

                        while (runsEnumerator.MoveNext())
                        {
                            XWPFRun run = (XWPFRun)runsEnumerator.Current;

                            if (run.GetCTR().alternateContent != null)
                            {

                                foreach (var alternateContent in run.GetCTR().alternateContent)
                                {
                                    string text = Utils.VerifyTextForXML(wordOutput.WordText.Text, "Placeholder:'" + wordOutput.Placeholder + "'", wordFile.Options.AutoRemoveInvalidXMLChars);
                                    WordTextBox.ReplaceTextInTextBox(alternateContent, wordOutput.Placeholder, text);
                                }

                            }
                        }
                    }
                }
            }
        }

        private static void ProcessTables(IEnumerator tablesEnumerator, WordStructures.WordFile wordFile, List<XWPFTable> TableToRemove)
        {
            while (tablesEnumerator.MoveNext())
            {
                XWPFTable table = (XWPFTable)tablesEnumerator.Current;
                ProcessTable(table, wordFile, TableToRemove);
            }
        }

        private static void ProcessTable(XWPFTable table, WordStructures.WordFile wordFile, List<XWPFTable> TableToRemove)
        {
            IEnumerator rows = table.Rows.GetEnumerator();
            while (rows.MoveNext())
            {
                List<XWPFParagraph> ParagraphToRemove = new List<XWPFParagraph>();
                XWPFTableRow row = (XWPFTableRow)rows.Current;
                IEnumerator cells = row.GetTableCells().GetEnumerator();
                while (cells.MoveNext())
                {
                    XWPFTableCell cell = (XWPFTableCell)cells.Current;
                    IEnumerator cellParagraphs = cell.Paragraphs.GetEnumerator();

                    ProcessParagraphs(cellParagraphs, wordFile, ParagraphToRemove);
                    if (ParagraphToRemove.Count != 0)
                    {
                        foreach (XWPFParagraph p in ParagraphToRemove)
                        {
                            if (cell.Paragraphs.Count > ParagraphToRemove.Count)
                            {
                                cell.GetCTTc().RemoveP(cell.Paragraphs.IndexOf(p));
                            }
                        }
                        ParagraphToRemove.Clear();
                    }
                }
            }

            ProcessWordTableOutputs(wordFile, table, TableToRemove);
        }



        private static void ProcessWordTableOutputs(WordStructures.WordFile wordFile, XWPFTable table, List<XWPFTable> TablesToRemove)
        {
            for (int i = 0; i < wordFile.WordOutputs.Count; i++)
            {
                WordStructures.WordOutput wordOutput = wordFile.WordOutputs[i];

                if (wordOutput.OutputType == (int)WordOutputType.LegacyTable || wordOutput.OutputType == (int)WordOutputType.Table)
                {
                    if (table.GetRow(0).GetCell(0) == null && wordOutput.OutputType == (int)WordOutputType.LegacyTable) //Remove if the style 2 because this || (table.GetRow(ssWordFile.ssSTWordFile.ssWordOutputs[i].ssSTWordOutput.ssWordCustomTable.ssSTWorldCustomTable.ssStartRow).GetCell(0) == null && ssWordFile.ssSTWordFile.ssWordOutputs[i].ssSTWordOutput.ssOutputType == 4))
                    {
                        new Exception("Template doesn't have a predefined table for replacement");
                    }
                    if ((table.GetRow(0).GetCell(0).GetText().Equals(wordOutput.Placeholder) && wordOutput.OutputType == (int)WordOutputType.LegacyTable))
                    {
                        if (wordOutput.DeletePlaceholder)
                        {
                            TablesToRemove.Add(table);
                        }
                        else
                        {
                            ProcessWordLegacyTable(wordOutput.WordLegacyTable, table, wordFile.Options.AutoRemoveInvalidXMLChars);
                        }
                    }
                    else
                    {
                        if (wordOutput.OutputType == (int)WordOutputType.Table)
                            if (table.Rows.Count >= wordOutput.WordTable.StartRow + 1)
                            {
                                if (table.GetRow(wordOutput.WordTable.StartRow).GetCell(0).GetText().IndexOf(wordOutput.Placeholder) > -1)
                                {
                                    ProcessWordTable(wordOutput.WordTable, table, wordFile.Options.AutoRemoveInvalidXMLChars);
                                }
                            }
                    }
                }
            }
        }


        private static void ProcessWordTable(WordStructures.WordTable wordTable, XWPFTable table, bool autoRemoveInvalidXMLChars)
        {

            int templateRowIndex = wordTable.StartRow;
            XWPFTableRow templateRow = table.GetRow(templateRowIndex); //Get template row

            List<XWPFTableRow> dummyRowsToAddBack = new List<XWPFTableRow>();

            bool hasDummyRows = templateRowIndex < table.NumberOfRows - 1;

            if (hasDummyRows)
            {
                int startIndexOfDummyRowsToKeep = templateRowIndex + 1 + wordTable.TableRows.Count - 1;

                int numberOfTableRows = table.NumberOfRows;

                for (int i = startIndexOfDummyRowsToKeep; i < numberOfTableRows; i++)
                {
                    dummyRowsToAddBack.Add(table.GetRow(i));
                }

                for (int i = templateRowIndex + 1; i < numberOfTableRows; i++) //Removing dummy rows
                {
                    table.RemoveRow(templateRowIndex + 1);
                }
            }

            for (int i = 0; i < wordTable.TableRows.Count; i++) //Creating new rows based on the template
            {
                WordStructures.WordTableRow tableRow = wordTable.TableRows[i];

                XWPFTableRow newRow = templateRow.CloneRow();

                IEnumerator tableCellsEnumerator = newRow.GetTableCells().GetEnumerator();

                while (tableCellsEnumerator.MoveNext())
                {
                    XWPFTableCell newTableCell = (XWPFTableCell)tableCellsEnumerator.Current;

                    for (int k = 0; k < tableRow.RowReplacements.Count; k++)
                    {
                        WordStructures.WordTableRowReplacement tableRowReplacement = tableRow.RowReplacements[k];

                        foreach (XWPFParagraph paragraph in newTableCell.Paragraphs)
                        {

                            int matchIndex = HasMatch(paragraph, tableRowReplacement.Placeholder);
                            if (matchIndex >= 0)
                            {
                                if (tableRowReplacement.WordPicture.Picture.Length == 0)
                                {
                                    tableRowReplacement.WordText.Text = Utils.VerifyTextForXML(tableRowReplacement.WordText.Text, ("Table:(Placeholder:'" + tableRowReplacement.Placeholder + "',RowIndex:" + k + ")"), autoRemoveInvalidXMLChars);
                                    ReplaceTextWithText(paragraph, tableRowReplacement.Placeholder, matchIndex, tableRowReplacement.WordText);
                                }
                                else if (tableRowReplacement.WordPicture.Picture.Length > 0)
                                {
                                    ReplaceTextWithPicture(paragraph, tableRowReplacement.Placeholder, matchIndex, tableRowReplacement.WordPicture);

                                }
                            }

                        }
                    }
                }
            }

            if (hasDummyRows)
            {
                foreach (XWPFTableRow row in dummyRowsToAddBack)
                {
                    table.AddRow(row);
                }
            }

            table.RemoveRow(templateRowIndex);
        }

        private static void ProcessWordLegacyTable(WordStructures.WordLegacyTable wordLegacyTable, XWPFTable table, bool autoRemoveInvalidXMLChars)
        {
            /*
             * For the each new created cell it's important to delete all paragraphs because by default .AddNewTableCell() and
             * .CreateRow() creates Cells with the default settings and while overriting them there are some issues.
             * 
             * There is a case when table have hight bigger than on the template - the reason is paragraph spacing after (<w:pPr> <w:spacing/>)
             * NPOI have some issues to set spacing in tables. If it's needed - on the Template set layout default spacing 0 or the value that is needed. 
            */
            XWPFTableCell templateCell = table.GetRow(0).GetCell(0);
            XWPFTableRow baseRow = table.GetRow(0);
            CT_Tc templateCellCT = templateCell.GetCTTc();
            XWPFParagraph templateParagraph = templateCell.Paragraphs[0];
            for (int i = 1; i < wordLegacyTable.TableRows[0].TableCells.Count; i++)
            {
                table.GetRow(0).AddNewTableCell();
            }

            for (int i = 1; i < wordLegacyTable.TableRows.Count; i++)
            {
                table.CreateRow();
                XWPFTableRow newRow = table.GetRow(i);
                newRow.Height = baseRow.Height;
            }

            for (int i = 0; i < wordLegacyTable.TableRows.Count; i++)
            {
                WordStructures.WordLegacyTableRow row = wordLegacyTable.TableRows[i];

                for (int j = 0; j < row.TableCells.Count; j++)
                {
                    XWPFTableCell tableCell = table.GetRow(i).GetCell(j);
                    WordStructures.WordLegacyTableCell cell = row.TableCells[j];
                    string text = Utils.VerifyTextForXML(cell.Value, ("LegacyTable:(RowIndex: " + i + ",ColumnIndex: " + j + ")"), autoRemoveInvalidXMLChars);

                    if (i == 0 && j == 0)
                    {
                        XWPFRun run = tableCell.Paragraphs[0].Runs[0];
                        CT_R r = run.GetCTR();
                        r.rPr = templateParagraph.Runs[0].GetCTR().rPr;
                        CT_Text textValue = (CT_Text)r.Items[0];
                        textValue.Value = "";


                        string[] newTextValue = GetTextLines(text);
                        if (newTextValue.Length > 1)
                        {
                            run.SetText(newTextValue[0], 0);

                            for (int k = 1; k != newTextValue.Length; k++)
                            {
                                run.AddBreak(BreakClear.ALL);
                                run.SetText(newTextValue[k], k);
                            }
                        }
                        else
                        {
                            run.SetText(text);
                        }

                        if (tableCell.Paragraphs[0].Runs.Count > 1)
                        {
                            for (int c = 1; c < tableCell.Paragraphs[0].Runs.Count; c++)
                            {
                                tableCell.Paragraphs[0].Runs[c].SetText("");
                            }
                        }
                    }
                    else
                    {
                        int a = 0;
                        while (a < tableCell.Paragraphs.Count)
                        {
                            tableCell.Paragraphs.RemoveAt(0);
                            a++;
                        }
                        tableCell.GetCTTc().tcPr = templateCellCT.tcPr;
                        tableCell.GetCTTc().Items.Clear();
                        foreach (var item in templateCellCT.Items)
                        {
                            if (item is CT_P) { } //old paragraph is not needed, only it's settings.
                            else
                            {
                                tableCell.GetCTTc().Items.Add(item);
                            }
                        }
                        tableCell.AddParagraph();
                        tableCell.Paragraphs[0].CreateRun();
                        XWPFRun run = tableCell.Paragraphs[0].Runs[0];
                        CT_R r = run.GetCTR();
                        r.rPr = templateParagraph.Runs[0].GetCTR().rPr;
                        r.AddNewInstrText();
                        CT_Text textValue = (CT_Text)r.Items[0];
                        textValue.Value = "";

                        string[] newTextValue = GetTextLines(text);
                        if (newTextValue.Length > 1)
                        {
                            run.SetText(newTextValue[0], 0);

                            for (int k = 1; k != newTextValue.Length; k++)
                            {
                                run.AddBreak(BreakClear.ALL);
                                run.SetText(newTextValue[k], k);
                            }
                        }
                        else
                        {
                            run.SetText(text);
                        }

                    }
                }
            }
        }

        private static string[] GetTextLines(string text)
        {
            text = SanitizedNewLines(text);

            return text.Split(new string[] { ((char)10).ToString() }, StringSplitOptions.None);
        }

        //sanitize new lines: replace "\r\n" and "\r" with "\n"
        private static string SanitizedNewLines(string text)
        {
            string newline = ((char)10).ToString() + (char)13;
            text = text.Replace(newline, ((char)10).ToString());
            newline = ((char)13).ToString() + (char)10;
            text = text.Replace(newline, ((char)10).ToString());
            text = text.Replace((char)13, (char)10);

            return text;
        }

        private static int HasMatch(XWPFParagraph p, string pattern)
        {
            string text = p.ParagraphText;
            return text.IndexOf(pattern);

        }

        private class TextIndex
        {
            public string Text { get; set; }
            public int StartIndex { get; private set; }
            public int EndIndex { get { return StartIndex + Text.Length - 1; } }
            public XWPFRun TextRun { get; set; }


            public TextIndex(XWPFRun run, string t, int startIndex)
            {
                this.Text = t;
                this.StartIndex = startIndex;
                this.TextRun = run;
            }
        }

        private static List<TextIndex> GetTextIndexList(XWPFParagraph p)
        {
            List<TextIndex> texts = new List<TextIndex>();
            StringBuilder concat = new StringBuilder();
            IEnumerator runs = p.Runs.GetEnumerator();
            while (runs.MoveNext())
            {
                XWPFRun run = (XWPFRun)runs.Current;
                int startIndex = concat.Length;
                texts.Add(new TextIndex(run, run.Text, startIndex));
                concat.Append(run.Text);

            }

            return texts;
        }

        private static void ReplaceTextWithText(XWPFParagraph p, string pattern, int matchIndex, WordStructures.WordText wordText)
        {
            //Get all the runs
            List<TextIndex> texts = GetTextIndexList(p);

            //Get all the runs that have placeholder
            int startRunIndex = texts.IndexOf(texts.Find(x => x.StartIndex <= matchIndex && x.EndIndex >= matchIndex));
            int placeholderEndIndex = matchIndex + pattern.Length - 1;
            int endRunIndex = texts.IndexOf(texts.Find(x => x.StartIndex <= placeholderEndIndex && x.EndIndex >= placeholderEndIndex));
            List<TextIndex> placeholderRuns = texts.GetRange(startRunIndex, endRunIndex - startRunIndex + 1);

            //Get all the text for those
            string runsText = placeholderRuns.Select(i => i.Text).Aggregate((i, j) => i + j);

            //Remove all but the first run    
            for (int i = endRunIndex; i > startRunIndex; i--)
            {
                p.RemoveRun(i);
            }

            XWPFRun run; // Placeholder run
            string runText = "";

            if (wordText.Hyperlink != "")
            {
                run = placeholderRuns[0].TextRun;

                string[] sa = runsText.Split(new string[] { pattern }, StringSplitOptions.None);

                if (sa[0] != "")
                {
                    run.SetText(sa[0], 0);
                }

                XWPFHyperlinkRun hyperlinkRun = p.InsertNewHyperlinkRun(startRunIndex + 1, wordText.Hyperlink);
                hyperlinkRun.SetStyle("Hyperlink");

                if (sa[1] != "")
                {
                    XWPFRun runAfter = p.InsertNewRun((startRunIndex + 2));
                    CloneRunProperties(run, runAfter);
                    runAfter.SetText(sa[1], 0);
                }

                if (sa[0] == "")
                {
                    p.RemoveRun(startRunIndex);
                }

                run = hyperlinkRun;
                runText = SanitizedNewLines(wordText.Text);
            }
            else
            {
                run = placeholderRuns[0].TextRun;
                runText = runsText.Replace(pattern, SanitizedNewLines(wordText.Text));
                if (runText.StartsWith("\t"))
                {
                    runText = runText.Substring(1); // Don't repeat tabs. (TODO: If user want to replace with the text that starts with tab)
                }
            }

            //Populates the first run with the replacement text
            if (runText.Contains("\n"))
            {
                //if the text within the run had already new lines they need to be removed.
                for (int i = run.GetCTR().SizeOfBrArray() - 1; i >= 0; i--)
                {
                    run.GetCTR().RemoveBr(i);
                }

                //remove all text tags except the first one
                for (int i = run.GetCTR().SizeOfTArray() - 1; i >= 1; i--)
                {
                    run.GetCTR().RemoveT(i);
                }

                string[] splitText = runText.Split(new string[] { "\n" }, StringSplitOptions.None);

                run.SetText(splitText[0], 0);

                for (int i = 1; i < splitText.Length; i++)
                {
                    run.AddBreak(BreakClear.ALL);
                    run.SetText(splitText[i], i);
                }

            }
            else
            {
                run.SetText(runText, 0);
            }

            ApplyStyleToRun(run, wordText.TextStyle);

            int nextMatchIndex = HasMatch(p, pattern);
            if (nextMatchIndex >= 0)
            {
                ReplaceTextWithText(p, pattern, nextMatchIndex, wordText);
            }
        }

        private static void ReplaceTextWithPicture(XWPFParagraph p, string pattern, int matchIndex, WordStructures.WordPicture wordPicture)
        {
            //Get all the runs
            List<TextIndex> texts = GetTextIndexList(p);

            //Get all the runs that have placeholder
            int startRunIndex = texts.IndexOf(texts.Find(x => x.StartIndex <= matchIndex && x.EndIndex >= matchIndex));
            int placeholderEndIndex = matchIndex + pattern.Length - 1;
            int endRunIndex = texts.IndexOf(texts.Find(x => x.StartIndex <= placeholderEndIndex && x.EndIndex >= placeholderEndIndex));
            List<TextIndex> placeholderRuns = texts.GetRange(startRunIndex, endRunIndex - startRunIndex + 1);

            //Get all the text for those
            string runsText = placeholderRuns.Select(i => i.Text).Aggregate((i, j) => i + j);

            //Remove all but the first run    
            for (int i = endRunIndex; i > startRunIndex; i--)
            {
                try
                {
                    p.RemoveRun(i);
                }
                catch (ArgumentException)
                {
                    throw new Exception("Cannot set hyperlink. The placeholder should be a regular text.");
                }
            }

            XWPFRun run = placeholderRuns[0].TextRun;
            run.SetText("", 0);

            string[] sa = runsText.Split(new string[] { pattern }, StringSplitOptions.None);

            if (sa[0] != "")
            {
                XWPFRun runBefore = p.InsertNewRun(startRunIndex);
                CloneRunProperties(run, runBefore);
                runBefore.SetText(sa[0], 0);
            }

            if (sa[1] != "")
            {
                XWPFRun runAfter;
                if (sa[0] != "")
                {
                    runAfter = p.InsertNewRun((startRunIndex + 2));
                }
                else
                {
                    runAfter = p.InsertNewRun((startRunIndex + 1));
                }
                CloneRunProperties(run, runAfter);
                runAfter.SetText(sa[1], 0);

            }

            AddPicture(run, wordPicture.Picture, wordPicture.Width);

            int nextMatchIndex = HasMatch(p, pattern);
            if (nextMatchIndex >= 0)
            {
                ReplaceTextWithPicture(p, pattern, nextMatchIndex, wordPicture);
            }
        }

        private static void CloneRunProperties(XWPFRun source, XWPFRun dest)
        { // clones the underlying w:rPr element
            CT_RPr rPrSource = source.GetCTR().rPr;
            if (rPrSource != null)
            {
                dest.GetCTR().rPr = rPrSource.Copy();
            }
        }

        private static void ApplyStyleToRun(XWPFRun run, WordStructures.WordTextStyle wordTextStyle)
        {
            if (wordTextStyle.Color != "")
            {
                run.SetColor(wordTextStyle.Color);
            }
            if (wordTextStyle.FontSize > 0)
            {
                run.FontSize = wordTextStyle.FontSize;
            }
            if (wordTextStyle.IsBold)
            {
                run.IsBold = true;
            }
            if (wordTextStyle.IsItalic)
            {
                run.IsItalic = true;
            }
            if (wordTextStyle.IsUnderlined)
            {
                run.Underline = UnderlinePatterns.Single;
            }
        }

        private static void AddPicture(XWPFRun run, byte[] picture, int pictureWidth)
        {

            Image img = Image.Load(picture);
            int[] dpi = ImageUtils.GetResolution(img);

            int width = ConvertPixelsToEmu(pictureWidth == 0 ? img.Width : pictureWidth, (float)dpi[0]);
            int height = ConvertPixelsToEmu((pictureWidth == 0 ? img.Height : (pictureWidth * img.Height) / img.Width), (float)dpi[1]);

            MemoryStream ms = new MemoryStream(picture);
            NPOI.SS.UserModel.PictureType pictureType = Utils.GetPictureType(picture);

            run.AddPicture(
                ms,
                (int)pictureType,
                "picture",
                width,
                height
            );
        }

        private static int ConvertPixelsToEmu(int pixels, float dpi)
        {
            float inch = (float)pixels / dpi;
            float EMU = 360000 * inch * (float)2.54;
            return (int)EMU;
        }

        public static string ExtractWordFileContent(byte[] ssWordBinary)
        {
            MemoryStream memoryStream = new MemoryStream(ssWordBinary);
            XWPFDocument document = new XWPFDocument(memoryStream);
            XWPFWordExtractor extractor = new XWPFWordExtractor(document);
            extractor.SetFetchHyperlinks(true);
            return extractor.Text;
        }
    }
}
