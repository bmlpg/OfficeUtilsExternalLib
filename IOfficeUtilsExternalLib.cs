using OutSystems.ExternalLibraries.SDK;

namespace OfficeUtilsExternalLib
{
    [OSInterface(Description = "Provides advanced export capabilities to Excel and Word. Part of OfficeUtils component, and not meant to be consumed directly.", IconResourceName = "OfficeUtilsExternalLib.resources.officeutils_logo.png",  Name = "OfficeUtilsExternalLib")]
    public interface IOfficeUtilsExternalLib
    {
        [OSAction(Description = "Generate Excel binary content, based on all operations specified in ExcelFile variable.")]
        public void GenerateExcelFile(
            [OSParameter(Description = "Holds information about template, and the list of all operations to be performed.")]
            ExcelStructures.ExcelFile ExcelFile,
            [OSParameter(Description = "Excel binary content generated, based on all operations specified in ExcelFile variable.")]
            out byte[] ExcelBinary
        );

        [OSAction(Description = "Generate Word binary content, based on all operations specified in WordFile variable.")]
        public void GenerateWordFile(
            [OSParameter(Description = "Holds information about template, and the list of all operations to be performed.")]
            WordStructures.WordFile WordFile,
            [OSParameter(Description = "Word binary content generated, based on all operations specified in WordFile variable.")]
            out byte[] WordBinary
        );

        [OSAction(Description = "Extracts Excel spreadsheet content.")]
        public void ExtractExcelFileContent(
            [OSParameter(Description = "Binary content of an Excel spreadsheet file.")]
            byte[] ExcelBinary,
            [OSParameter(Description = "Excel file content in text.")]
            out string Content
        );

        [OSAction(Description = "Extracts Word document content.")]
        public void ExtractWordFileContent(
            [OSParameter(Description = "Binary content of an Word document file.")]
            byte[] WordBinary,
            [OSParameter(Description = "Word file content in text.")]
            out string Content
        );

        [OSAction(Description = "Run this action within a Timer if you want to prevent AWS Lambda \"cold starts\".")]
        public void Ping();
    }
}
