using NPOI.SS.UserModel;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats;
using SixLabors.ImageSharp.Metadata;

namespace OfficeUtilsExternalLib
{
    internal class Utils
    {
        //Copied from ImageUtils.cs, and customized
        public static int[] GetResolution(Image r)
        {
            ImageMetadata imageMetadata = r.Metadata;

            double horizontalResolution = 0;
            double verticalResolution = 0;

            if (imageMetadata.ResolutionUnits == PixelResolutionUnit.PixelsPerMeter)
            {
                horizontalResolution = imageMetadata.HorizontalResolution * 0.0254D;
                verticalResolution = imageMetadata.VerticalResolution * 0.0254D;
            }
            else if (imageMetadata.ResolutionUnits == PixelResolutionUnit.PixelsPerCentimeter)
            {
                horizontalResolution = imageMetadata.HorizontalResolution * 2.54D;
                verticalResolution = imageMetadata.VerticalResolution * 2.54D;
            }
            else
            {
                horizontalResolution = imageMetadata.HorizontalResolution;
                verticalResolution = imageMetadata.VerticalResolution;
            }

            return new int[] { (int)Math.Round(horizontalResolution), (int)Math.Round(verticalResolution) };
        }

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
    }
}
