namespace OfficeUtilsExternalLib
{
    public class OfficeUtilsExternalLib : IOfficeUtilsExternalLib
    {
        public void GenerateExcelFile(ExcelStructures.ExcelFile ExcelFile, out byte[] ExcelBinary)
        {
            ExcelBinary = Excel.GenerateExcelFile(ExcelFile);
        }

        public void GenerateWordFile(WordStructures.WordFile WordFile, out byte[] WordBinary)
        {
            WordBinary = Word.GenerateWordFile(WordFile);
        }

        public void  ExtractExcelFileContent(byte[] ExcelBinary, out string Content)
        {
            Content = Excel.ExtractExcelFileContent(ExcelBinary);
        }

        public void ExtractWordFileContent(byte[] WordBinary, out string Content)
        {
            Content = Word.ExtractWordFileContent(WordBinary);
        }

        public void Ping()
        {

        }
    }
}
