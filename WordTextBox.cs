using System.Collections;
using NPOI.OpenXmlFormats.Vml;
using NPOI.XWPF.UserModel;
using System.Xml;

namespace OfficeUtilsExternalLib
{
    internal class WordTextBox
    {
        public static void ReplaceTextInTextBox(CT_AlternateContent alternateContent, string pattern, string replaceText)
        {
            string textBoxXML = alternateContent.InnerXml;

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml("<root>" + textBoxXML + "</root>");

            XmlNamespaceManager nm = new XmlNamespaceManager(xmlDocument.NameTable);
            nm.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");


            XmlNodeList nodeList = xmlDocument.SelectNodes(".//*/w:txbxContent/w:p", nm);

            IEnumerator paragraphNodesEnumerator = nodeList.GetEnumerator();
            while (paragraphNodesEnumerator.MoveNext())
            {
                XmlNode paragraphNode = (XmlNode)paragraphNodesEnumerator.Current;
                ReplaceParagraphText(paragraphNode, nm, pattern, replaceText);
            }

            alternateContent.InnerXml = xmlDocument.FirstChild.InnerXml;
        }

        private static void ReplaceParagraphText(XmlNode paragraphNode, XmlNamespaceManager nm, string oldText, string newText)
        {
            XmlNodeList runNodes = paragraphNode.SelectNodes("w:r", nm);

            if (string.IsNullOrEmpty(oldText))
            {
                throw new ArgumentNullException("oldText should not be null");
            }
            TextSegment ts = SearchTextInParagraph(paragraphNode, nm, oldText, new PositionInParagraph() { Run = 0 });
            if (ts == null)
                return;
            if (ts.BeginRun == ts.EndRun)
            {
                ReplaceRunText(runNodes[ts.BeginRun], nm, oldText, newText);
            }
            else
            {
                ReplaceRunText(
                    runNodes[ts.BeginRun],
                    nm,
                    GetRunText(runNodes[ts.BeginRun], nm).Substring(ts.BeginChar),
                    newText + GetRunText(runNodes[ts.EndRun], nm).Substring(ts.EndChar + 1)
                );

                for (int i = ts.EndRun; i > ts.BeginRun; i--)
                {
                    paragraphNode.RemoveChild(runNodes[i]);
                }
            }
        }

        private static TextSegment SearchTextInParagraph(XmlNode paragraphNode, XmlNamespaceManager nm, String searched, PositionInParagraph startPos)
        {
            XmlNodeList runNodes = paragraphNode.SelectNodes("w:r", nm);

            int startRun = startPos.Run,
                startText = startPos.Text,
                startChar = startPos.Char;
            int beginRunPos = 0, beginTextPos = 0, beginCharPos = 0, candCharPos = 0;
            bool newList = false;
            for (int runPos = startRun; runPos < runNodes.Count; runPos++)
            {
                int textPos = 0, charPos = 0;
                XmlNode run = runNodes[runPos];
                XmlNodeList runItems = run.ChildNodes;
                foreach (XmlNode item in runItems)
                {
                    if (item.Name == "w:t")
                    {
                        if (textPos >= startText)
                        {
                            String candidate = item.FirstChild.InnerText;
                            if (runPos == startRun)
                                charPos = startChar;
                            else
                                charPos = 0;
                            for (; charPos < candidate.Length; charPos++)
                            {
                                if ((candidate[charPos] == searched[0]) && (candCharPos == 0))
                                {
                                    beginTextPos = textPos;
                                    beginCharPos = charPos;
                                    beginRunPos = runPos;
                                    newList = true;
                                }
                                if (candidate[charPos] == searched[candCharPos])
                                {
                                    if (candCharPos + 1 < searched.Length)
                                    {
                                        candCharPos++;
                                    }
                                    else if (newList)
                                    {
                                        TextSegment segment = new TextSegment();
                                        segment.BeginRun = (beginRunPos);
                                        segment.BeginText = (beginTextPos);
                                        segment.BeginChar = (beginCharPos);
                                        segment.EndRun = (runPos);
                                        segment.EndText = (textPos);
                                        segment.EndChar = (charPos);
                                        return segment;
                                    }
                                }
                                else
                                    candCharPos = 0;
                            }
                        }
                        textPos++;
                    }
                    else if (item.Name == "w:proofErr")
                    {
                        //c.RemoveXml();
                    }
                    else if (item.Name == "w:rPr")
                    {
                        //do nothing
                    }
                    else
                        candCharPos = 0;
                }
            }
            return null;
        }

        private static void ReplaceRunText(XmlNode runNode, XmlNamespaceManager nm, string oldText, string newText)
        {
            XmlNode textNode = runNode.SelectSingleNode("w:t", nm);
            if (textNode != null)
            {
                textNode.FirstChild.Value = textNode.FirstChild.Value.Replace(oldText, newText);
            }
        }

        private static string GetRunText(XmlNode runNode, XmlNamespaceManager nm)
        {
            XmlNode textNode = runNode.SelectSingleNode("w:t", nm);
            if (textNode != null)
            {
                return textNode.FirstChild.Value;
            }
            return "";
        }
    }
}
