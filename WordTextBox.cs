using System.Collections;
using System.Xml;
using NPOI.OpenXmlFormats.Vml;
using NPOI.XWPF.UserModel;
using OfficeUtilsExternalLib.WordStructures;

namespace OfficeUtilsExternalLib
{
    internal class WordTextBox
    {
        public static void ReplaceTextInTextBox(CT_AlternateContent alternateContent, string pattern, string replaceText, WordTextStyle textStyle)
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
                ReplaceParagraphText(paragraphNode, nm, pattern, replaceText, textStyle);
            }

            alternateContent.InnerXml = xmlDocument.FirstChild.InnerXml;
        }

        private static void ReplaceParagraphText(XmlNode paragraphNode, XmlNamespaceManager nm, string oldText, string newText, WordTextStyle textStyle)
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
                ReplaceRunText(
                    runNodes[ts.BeginRun],
                    nm,
                    oldText,
                    newText,
                    textStyle
                );
            }
            else
            {
                ReplaceRunText(
                    runNodes[ts.BeginRun],
                    nm,
                    GetRunText(runNodes[ts.BeginRun], nm).Substring(ts.BeginChar),
                    newText + GetRunText(runNodes[ts.EndRun], nm).Substring(ts.EndChar + 1),
                    textStyle
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
                            String candidate = item.InnerText;
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

        private static void ReplaceRunText(XmlNode runNode, XmlNamespaceManager nm, string oldText, string newText, WordTextStyle textStyle)
        {
            XmlNode textNode = runNode.SelectSingleNode("w:t", nm);
            if (textNode != null)
            {
                textNode.FirstChild.Value = textNode.FirstChild.Value.Replace(oldText, newText);
            }

            XmlNode propertiesNode = runNode.SelectSingleNode("w:rPr", nm);
            if (propertiesNode == null)
            {
                propertiesNode = runNode.OwnerDocument.CreateElement("rPr", nm.LookupNamespace("w"));
                runNode.PrependChild(propertiesNode);
            }

            if (textStyle.Color != "")
            {
                XmlElement colorElem = (XmlElement)(propertiesNode.SelectSingleNode("w:color", nm)
                       ?? propertiesNode.AppendChild(runNode.OwnerDocument.CreateElement("w", "color", nm.LookupNamespace("w"))));
                colorElem.SetAttribute("val", nm.LookupNamespace("w"), textStyle.Color);
            }
            if (textStyle.FontSize > 0)
            {
                XmlElement szElem = (XmlElement)(propertiesNode.SelectSingleNode("w:sz", nm)
                       ?? propertiesNode.AppendChild(runNode.OwnerDocument.CreateElement("w", "sz", nm.LookupNamespace("w"))));
                szElem.SetAttribute("val", nm.LookupNamespace("w"), (textStyle.FontSize * 2).ToString());

                XmlElement szCsElem = (XmlElement)(propertiesNode.SelectSingleNode("w:szCs", nm)
                       ?? propertiesNode.AppendChild(runNode.OwnerDocument.CreateElement("w", "szCs", nm.LookupNamespace("w"))));
                szCsElem.SetAttribute("val", nm.LookupNamespace("w"), (textStyle.FontSize * 2).ToString());
            }
            if (textStyle.IsBold)
            {
                XmlElement bElem = (XmlElement)(propertiesNode.SelectSingleNode("w:b", nm)
                       ?? propertiesNode.AppendChild(runNode.OwnerDocument.CreateElement("w", "b", nm.LookupNamespace("w"))));
                bElem.SetAttribute("val", nm.LookupNamespace("w"), "true");
            }
            if (textStyle.IsItalic)
            {
                XmlElement iElem = (XmlElement)(propertiesNode.SelectSingleNode("w:i", nm)
                       ?? propertiesNode.AppendChild(runNode.OwnerDocument.CreateElement("w", "i", nm.LookupNamespace("w"))));
                iElem.SetAttribute("val", nm.LookupNamespace("w"), "true");
            }
            if (textStyle.IsUnderlined)
            {
                XmlElement uElem = (XmlElement)(propertiesNode.SelectSingleNode("w:u", nm)
                       ?? propertiesNode.AppendChild(runNode.OwnerDocument.CreateElement("w", "u", nm.LookupNamespace("w"))));
                uElem.SetAttribute("val", nm.LookupNamespace("w"), "single");
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
