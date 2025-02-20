using SixLabors.ImageSharp;
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
    }
}
