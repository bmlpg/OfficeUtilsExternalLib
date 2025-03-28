using System.Collections;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.OpenXml4Net.OPC;
using System.Text;
using SixLabors.ImageSharp;
using NPOI.XWPF.Extractor;

namespace OfficeUtilsExternalLib
{
    internal class Word
    {

        public enum WordOutputType : int
        {
            Text = 1,
            TableLegacy = 2,
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

            ProcessParagraphs(documentParagraphsIterator, wordFile, ParagraphsToRemove, document);

            if (ParagraphsToRemove.Count != 0)
            {
                foreach (XWPFParagraph p in ParagraphsToRemove)
                {
                    document.RemoveBodyElement(document.GetPosOfParagraph(p));
                }
                ParagraphsToRemove.Clear();
            }

            IEnumerator documentTablesIterator = document.GetTablesEnumerator();
            ProcessTables(documentTablesIterator, wordFile, TablesToRemove, document);

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
                ProcessParagraphs(headerParagraphsEnumerator, wordFile, ParagraphsToRemove, header);

                IEnumerator headerTablesEnumerator = header.Tables.GetEnumerator();
                ProcessTables(headerTablesEnumerator, wordFile, TablesToRemove, header);
            }


            //Footer
            IEnumerator footerEnumerator = document.FooterList.GetEnumerator();

            while (footerEnumerator.MoveNext())
            {
                XWPFFooter footer = (XWPFFooter)footerEnumerator.Current;

                IEnumerator footerParagraphsEnumerator = footer.GetListParagraph().GetEnumerator();
                ProcessParagraphs(footerParagraphsEnumerator, wordFile, ParagraphsToRemove, footer);

                IEnumerator footerTablesEnumerator = footer.Tables.GetEnumerator();
                ProcessTables(footerTablesEnumerator, wordFile, TablesToRemove, footer);
            }

            MemoryStream outputStream = new MemoryStream();
            document.Write(outputStream);
            byte[] wordBinary = outputStream.ToArray();

            return wordBinary;

        } // MssGenerateWordFile

        private static void ProcessParagraphs(IEnumerator paragraphsEnumerator, WordStructures.WordFile wordFile, List<XWPFParagraph> ParagraphsToRemove, Object context)
        {
            while (paragraphsEnumerator.MoveNext())
            {
                XWPFParagraph paragraph = (XWPFParagraph)paragraphsEnumerator.Current;
                ProcessParagraph(paragraph, wordFile, ParagraphsToRemove, context);
            }
        }

        private static void ProcessParagraph(XWPFParagraph paragraph, WordStructures.WordFile wordFile, List<XWPFParagraph> ParagraphsToRemove, Object context)
        {
            for (int i = 0; i < wordFile.WordOutputs.Count; i++)
            {
                WordStructures.WordOutput wordOutput = wordFile.WordOutputs[i];

                if (wordOutput.OutputType == (int)WordOutputType.Text || wordOutput.OutputType == (int)WordOutputType.Picture)
                {
                    int matchSPIndex = HasMatch(paragraph, wordOutput.Placeholder);
                    if (matchSPIndex >= 0)
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
                                ReplaceText(paragraph, wordOutput.Placeholder, "", matchSPIndex);
                            }
                        }
                        else
                        {
                            if (wordOutput.OutputType == (int)WordOutputType.Text)
                            {
                                ReplaceText(paragraph, wordOutput.Placeholder, wordOutput.WordText.Text, matchSPIndex, wordOutput.WordText.Hyperlink);
                            }
                            else if (wordOutput.OutputType == (int)WordOutputType.Picture)
                            {
                                ProcessWordPicture(wordOutput.Placeholder, wordOutput.WordPicture, paragraph, context, matchSPIndex);
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
                                WordTextBox.ReplaceTextInTextBox(run.GetCTR().alternateContent, wordOutput.Placeholder, wordOutput.WordText.Text);
                            }
                        }
                    }
                }
            }
        }

        private static void ProcessTables(IEnumerator tablesEnumerator, WordStructures.WordFile wordFile, List<XWPFTable> TableToRemove, object context)
        {
            while (tablesEnumerator.MoveNext())
            {
                XWPFTable table = (XWPFTable)tablesEnumerator.Current;
                ProcessTable(table, wordFile, TableToRemove, context);
            }
        }

        private static void ProcessTable(XWPFTable table, WordStructures.WordFile wordFile, List<XWPFTable> TableToRemove, object context)
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

                    ProcessParagraphs(cellParagraphs, wordFile, ParagraphToRemove, context);
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

            ProcessWordTableOutputs(wordFile, table, TableToRemove, context);
        }



        private static void ProcessWordTableOutputs(WordStructures.WordFile wordFile, XWPFTable table, List<XWPFTable> TableToRemove, Object context)
        {
            for (int i = 0; i < wordFile.WordOutputs.Count; i++)
            {
                if (wordFile.WordOutputs[i].OutputType == (int)WordOutputType.TableLegacy || wordFile.WordOutputs[i].OutputType == (int)WordOutputType.Table)
                {
                    if (table.GetRow(0).GetCell(0) == null && wordFile.WordOutputs[i].OutputType == (int)WordOutputType.TableLegacy) //Remove if the style 2 because this || (table.GetRow(ssWordFile.ssSTWordFile.ssWordOutputs[i].ssSTWordOutput.ssWordCustomTable.ssSTWorldCustomTable.ssStartRow).GetCell(0) == null && ssWordFile.ssSTWordFile.ssWordOutputs[i].ssSTWordOutput.ssOutputType == 4))
                    {
                        new Exception("Template doesn't have a predefined table for replacement");
                    }
                    if ((table.GetRow(0).GetCell(0).GetText().Equals(wordFile.WordOutputs[i].Placeholder) && wordFile.WordOutputs[i].OutputType == (int)WordOutputType.TableLegacy))
                    {
                        if (wordFile.WordOutputs[i].DeletePlaceholder)
                        {
                            TableToRemove.Add(table);
                        }
                        else
                        {
                            ProcessWordLegacyTable(wordFile.WordOutputs[i].WordLegacyTable, table);
                        }
                    }
                    else
                    {
                        if (wordFile.WordOutputs[i].OutputType == (int)WordOutputType.Table)
                            if (table.Rows.Count >= wordFile.WordOutputs[i].WordTable.StartRow + 1)
                            {
                                if (table.GetRow(wordFile.WordOutputs[i].WordTable.StartRow).GetCell(0).GetText().IndexOf(wordFile.WordOutputs[i].Placeholder) > -1)
                                {
                                    ProcessWordTable(wordFile.WordOutputs[i].WordTable, table, context);
                                }
                            }
                    }
                }
            }
        }


        private static void ProcessWordTable(WordStructures.WordTable wordTable, XWPFTable table, Object context)
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

                //XWPFTableRow newRow = templateRow.CloneRow();
                XWPFTableRow newRow = CloneRow(templateRow);

                IEnumerator tableCellsEnumerator = newRow.GetTableCells().GetEnumerator();

                while (tableCellsEnumerator.MoveNext())
                {
                    XWPFTableCell newTableCell = (XWPFTableCell)tableCellsEnumerator.Current;

                    for (int k = 0; k < tableRow.RowReplacements.Count; k++)
                    {
                        WordStructures.WordTableRowReplacement tableRowReplacement = tableRow.RowReplacements[k];

                        foreach (XWPFParagraph paragr in newTableCell.Paragraphs)
                        {   
                            int matchSPIndex = HasMatch(paragr, tableRowReplacement.Placeholder);
                            if (matchSPIndex >= 0)
                            {

                                if (tableRowReplacement.Picture.Length == 0)
                                {
                                    ReplaceText(paragr, tableRowReplacement.Placeholder, tableRowReplacement.Text, matchSPIndex);
                                }
                                else if (tableRowReplacement.Picture.Length > 0)
                                {
                                    WordStructures.WordPicture pic;
                                    pic.Picture = tableRowReplacement.Picture;
                                    pic.Width = tableRowReplacement.PictureWidth;
                                    ProcessWordPicture(tableRowReplacement.Placeholder, pic, paragr, context, matchSPIndex);
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

        private static XWPFTableRow CloneRow(XWPFTableRow templateRow)
        {
            XWPFTable table = templateRow.GetTable();
            
            XWPFTableRow newRow = table.CreateRow();
            newRow.Height = templateRow.Height;
            newRow.IsCantSplitRow = templateRow.IsCantSplitRow;
            newRow.IsRepeatHeader = templateRow.IsRepeatHeader;
            for (int j = 0; j < templateRow.GetTableCells().Count; j++) // For each cell in the row
            {
                XWPFTableCell newCell = newRow.GetCell(j);
                newCell.GetCTTc().tcPr = templateRow.GetCell(j).GetCTTc().tcPr; //Copy the properties of the cell
                newCell.RemoveParagraph(0); //new cell is created with default stuff, that we don't want, we need a template copy
                newCell.GetCTTc().Items.Clear();
                ArrayList cT_TcItems = (ArrayList)templateRow.GetCell(j).GetCTTc().Items.Clone();
                foreach (var item in cT_TcItems) // Add all items from the template
                {

                    if (item is CT_P p)
                    {
                        XWPFParagraph nP = newCell.AddParagraph();

                        CT_P newP = newCell.GetCTTc().GetPList().Last();

                        if (!p.pPr.IsEmpty)
                        {
                            newP.pPr = p.pPr;
                        }

                        foreach (CT_R run in p.GetRList())
                        {

                            XWPFRun nR = nP.CreateRun();
                            nR.GetCTR().rPr = run.rPr;
                            if (run.Items.Count > 0)
                            {
                                for (int c = 0; c < run.Items.Count; c++)
                                {
                                    if (run.Items[c] is CT_Text)
                                    {
                                        nR.SetText(run.GetTList().First().Value);
                                    }
                                    else if (run.Items[c] is CT_Empty && run.ItemsElementName[c] == RunItemsChoiceType.tab)
                                    {
                                        nR.AddTab();
                                    }
                                }
                            }


                        }
                    }
                    else newCell.GetCTTc().Items.Add(item);


                }
            }

            return newRow;
        }

        private static void ProcessWordLegacyTable(WordStructures.WordLegacyTable wordLegacyTable, XWPFTable table)
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
                    if (i == 0 && j == 0)
                    {
                        XWPFRun run = tableCell.Paragraphs[0].Runs[0];
                        CT_R r = run.GetCTR();
                        r.rPr = templateParagraph.Runs[0].GetCTR().rPr;
                        CT_Text textValue = (CT_Text)r.Items[0];
                        textValue.Value = "";

                        string[] newTextValue = ProcessNewLines(cell.Value);
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
                            run.SetText(cell.Value);
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
                        string[] newTextValue = ProcessNewLines(cell.Value);
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
                            run.SetText(cell.Value);
                        }

                    }
                }
            }
        }

        private static string[] ProcessNewLines(string originalText)
        {
            string newline = ((char)10).ToString() + (char)13;
            originalText = originalText.Replace(newline, ((char)10).ToString());
            newline = ((char)13).ToString() + (char)10;
            originalText = originalText.Replace(newline, ((char)10).ToString());
            originalText = originalText.Replace((char)13, (char)10);

            return originalText.Split(new string[] { ((char)10).ToString() }, StringSplitOptions.None);
        }

        private static int HasMatch(XWPFParagraph p, string pattern)
        {

            string text = GetTexts(p);
            return text.IndexOf(pattern);

        }

        private static string GetTexts(XWPFParagraph p)
        {
            StringBuilder concat = new StringBuilder();
            IEnumerator runs = p.Runs.GetEnumerator();
            while (runs.MoveNext())
            {
                XWPFRun run = (XWPFRun)runs.Current;
                StringBuilder text = new StringBuilder();
                for (int i = 0; i < run.GetCTR().Items.Count; i++)
                {
                    object o = run.GetCTR().Items[i];
                    if (o is CT_Text)
                    {

                        if (!(run.GetCTR().ItemsElementName[i] == RunItemsChoiceType.instrText))
                        {
                            text.Append(((CT_Text)o).Value);
                        }
                    }

                    // Complex type evaluation (currently only for extraction of check boxes)
                    if (o is CT_FldChar)
                    {
                        CT_FldChar ctfldChar = ((CT_FldChar)o);
                        if (ctfldChar.fldCharType == ST_FldCharType.begin)
                        {
                            if (ctfldChar.ffData != null)
                            {
                                foreach (CT_FFCheckBox checkBox in ctfldChar.ffData.GetCheckBoxList())
                                {
                                    if (checkBox.@default.val == true)
                                    {
                                        text.Append("|X|");
                                    }
                                    else
                                    {
                                        text.Append("|_|");
                                    }
                                }
                            }
                        }
                    }

                    if (o is CT_PTab)
                    {
                        text.Append("\t");
                    }
                    if (o is CT_Br)
                    {
                        text.Append("\n");
                    }

                    if (o is CT_Empty)
                    {
                        // Some inline text elements Get returned not as
                        //  themselves, but as CTEmpty, owing to some odd
                        //  defInitions around line 5642 of the XSDs
                        // This bit works around it, and replicates the above
                        //  rules for that case
                        if (run.GetCTR().ItemsElementName[i] == RunItemsChoiceType.tab)
                        {
                            text.Append("\t");
                        }
                        if (run.GetCTR().ItemsElementName[i] == RunItemsChoiceType.br)
                        {
                            text.Append("\n");
                        }
                        if (run.GetCTR().ItemsElementName[i] == RunItemsChoiceType.cr)
                        {
                            text.Append("\n");
                        }
                    }
                    if (o is CT_FtnEdnRef)
                    {
                        CT_FtnEdnRef ftn = (CT_FtnEdnRef)o;
                        String footnoteRef = ftn.DomNode.LocalName.Equals("footnoteReference") ?
                            "[footnoteRef:" + ftn.id + "]" : "[endnoteRef:" + ftn.id + "]";
                        text.Append(footnoteRef);
                    }
                }


                concat.Append(text.ToString());

            }
            return concat.ToString();
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

        private static void ReplaceText(XWPFParagraph p, string pattern, string replaceText, int matchIndex, string url = "")
        {
            //sanitize new lines: replace "\r\n" and "\r" with "\n"
            string newline = ((char)10).ToString() + (char)13;
            replaceText = replaceText.Replace(newline, ((char)10).ToString());
            newline = ((char)13).ToString() + (char)10;
            replaceText = replaceText.Replace(newline, ((char)10).ToString());
            replaceText = replaceText.Replace((char)13, (char)10);

            //Get all the runs
            List<TextIndex> texts = GetTextIndexList(p);

            //Get all the runs that have placeholder
            int startRunIndex = texts.IndexOf(texts.Find(x => x.StartIndex <= matchIndex && x.EndIndex >= matchIndex));
            int placeholderEndIndex = matchIndex + pattern.Length - 1;
            int endRunIndex = texts.IndexOf(texts.Find(x => x.StartIndex <= placeholderEndIndex && x.EndIndex >= placeholderEndIndex));
            List<TextIndex> placeholderRuns = texts.GetRange(startRunIndex, endRunIndex - startRunIndex + 1);

            //Get all the text for those runs and replace it
            string runsText = placeholderRuns.Select(i => i.Text).Aggregate((i, j) => i + j);
            string newRunText = runsText.Replace(pattern, replaceText);
            if (newRunText.StartsWith("\t"))
            {
                newRunText = newRunText.Substring(1); // Don't repeat tabs. (TODO: If user want to replace with the text that starts with tab)
            }

            XWPFRun startRun = texts[startRunIndex].TextRun;

            PackageRelationship newRelationship = null;
            if (url != "") // This is if URL needs to be changed under the pattern
            {
                if (startRun is XWPFHyperlinkRun)
                {
                    foreach (var part in p.Document.Package.GetParts())
                    {
                        PackagePart PP = (PackagePart)part;
                        if (PP.PartName.Name == "/word/document.xml")
                        {
                            newRelationship = PP.AddExternalRelationship(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
                        }
                    }

                    if (newRelationship == null)
                        new Exception("Can't create new URL reference");
                }
                else
                {
                    bool urlFound = false;
                    for (int i = startRunIndex; i >= 0; i--)
                    {
                        if (texts[i].TextRun.GetCTR().Items.Count > 0)
                        {
                            if (texts[i].TextRun.GetCTR().Items[0] is CT_Text text)
                            {
                                if (text.Value.Contains("HYPERLINK"))
                                {
                                    urlFound = true;
                                    text.Value = "HYPERLINK " + url;
                                }
                            }
                        }
                    }
                    if (!urlFound)
                        new Exception("URL is not created on the template");
                }
            }


            //Populate with replaced text first run
            if (newRunText.Contains("\n"))
            {
                //if the text within the run had already new lines than they need to be removed, because the logic is
                //to recreate them again (because the can be new lines in the replacement text and NPOI can add <br> only to the end of the list)

                string[] splitText = newRunText.Split(new string[] { "\n" }, StringSplitOptions.None);
                startRun.SetText(splitText[0], 0);

                foreach (var item in startRun.GetCTR().Items)
                {
                    if (item is CT_Br)
                    {
                        startRun.GetCTR().Items.Remove(item);
                    }
                }

                for (int i = 1; i != splitText.Length; i++)
                {
                    startRun.AddBreak(BreakClear.ALL);
                    startRun.SetText(splitText[i], i);
                }

            }
            else
            {
                startRun.SetText(newRunText, 0);
            }

            if (startRun is XWPFHyperlinkRun run1 && newRelationship != null)
            {
                run1.GetCTHyperlink().id = newRelationship.Id;
            }

            //Empty the next runs
            for (int i = startRunIndex + 1; i <= endRunIndex; i++)
            {
                XWPFRun nextRun = texts[i].TextRun;
                nextRun.SetText("");
                if (nextRun is XWPFHyperlinkRun run2 && newRelationship != null)
                {
                    run2.GetCTHyperlink().id = newRelationship.Id;
                }
            }

            int matchSPIndex = HasMatch(p, pattern);
            if (matchSPIndex >= 0)
            {
                ReplaceText(p, pattern, replaceText, matchSPIndex, url);
            }
        }

        private static void ProcessWordPicture(string placeholder, WordStructures.WordPicture wordPicture, XWPFParagraph paragraph, Object context, int matchIndex)
        {
            ReplaceText(paragraph, placeholder, "", matchIndex);
            IEnumerator runs = paragraph.Runs.GetEnumerator();
            if (runs.MoveNext())
            {
                XWPFRun run = (XWPFRun)runs.Current;
                while (runs.MoveNext())
                {
                    run = (XWPFRun)runs.Current;
                }

                Image img = Image.Load(wordPicture.Picture);
                int[] dpi = Utils.GetResolution(img);

                int picWidth = wordPicture.Width;
                int width = ConvertPixelsToEmu(picWidth == 0 ? img.Width : picWidth, (float)dpi[0]);
                int height = ConvertPixelsToEmu((picWidth == 0 ? img.Height : (picWidth * img.Height) / img.Width), (float)dpi[1]);

                MemoryStream ms = new MemoryStream(wordPicture.Picture);
                NPOI.SS.UserModel.PictureType pictureType = Utils.GetPictureType(wordPicture.Picture);

                run.AddPicture(
                    ms,
                    (int)pictureType,
                    "picture",
                    width,
                    height
                );
            }
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
