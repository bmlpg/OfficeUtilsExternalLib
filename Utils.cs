using System.Xml;
using NPOI.SS.UserModel;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats;

namespace OfficeUtilsExternalLib
{
    internal class Utils
    {
        public static PictureType GetPictureType(byte[] pictureBinary)
        {
            IImageFormat imageFormat = Image.DetectFormat(pictureBinary);

            PictureType pictureType = new PictureType();

            switch (imageFormat.Name)
            {
                case "EMF":
                    pictureType = PictureType.EMF;
                    break;
                case "WMF":
                    pictureType = PictureType.WMF;
                    break;
                case "PICT":
                    pictureType = PictureType.PICT;
                    break;
                case "JPEG":
                    pictureType = PictureType.JPEG;
                    break;
                case "PNG":
                    pictureType = PictureType.PNG;
                    break;
                case "DIB":
                    pictureType = PictureType.DIB;
                    break;
                case "GIF":
                    pictureType = PictureType.GIF;
                    break;
                case "TIFF":
                    pictureType = PictureType.TIFF;
                    break;
                case "EPS":
                    pictureType = PictureType.EPS;
                    break;
                case "BMP":
                    pictureType = PictureType.BMP;
                    break;
                case "WPG":
                    pictureType = PictureType.WPG;
                    break;
                default:
                    pictureType = PictureType.Unknown;
                    break;
            }

            return pictureType;
        }

        public static string VerifyTextForXML(string text, string context, bool autoRemoveInvalidXMLChars)
        {
            try
            {
                return XmlConvert.VerifyXmlChars(text);
            }
            catch (XmlException e)
            {
                if (autoRemoveInvalidXMLChars)
                {
                    char[] validXmlChars = text.Where(ch => XmlConvert.IsXmlChar(ch)).ToArray();
                    return new string(validXmlChars); ;
                }
                else
                {
                    int position = text.Select((value, index) => new { value, index }).Where(ch => !XmlConvert.IsXmlChar(ch.value)).Select(ch => ch.index).DefaultIfEmpty(-1).FirstOrDefault();
                    throw new Exception(
                        e.Message + " | " +
                        "Text:'" + text.Substring(0,50) + (text.Length > 50 ? "..." : "") + "'" + " | " +
                        "Position:" + position + " | " +
                        "Context:(" + context + ")" + " | " +
                        "Fix your input or set option 'AutoRemoveInvalidXMLChars' to 'true' to remove invalid XML characters automatically."
                    );
                }

            }
        }
    }
}
