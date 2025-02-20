using OutSystems.ExternalLibraries.SDK;

namespace OfficeUtilsExternalLib.ExcelStructures
{
    [OSStructure(Description = "Contains template spreadsheett, and a list of all operations to be performed.")]
    public struct ExcelFile
    {
        [OSStructureField(Description = "ExcelOutput record list", IsMandatory = true)]
        public List<ExcelOutput> ExcelOutputs;
        [OSStructureField(DataType = OSDataType.BinaryData, Description = "Template excel file", IsMandatory = false)]
        public byte[] Template;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "If True, locks the structure", IsMandatory = false)]
        public bool LockStructure;
    }

    [OSStructure(Description = "Operations to be performed to a specified sheet.")]
    public struct ExcelOutput
    {
        [OSStructureField(DataType = OSDataType.Integer, Description = "1 - table, 2 - cell, 3 - picture, 4 - merge cells, 5 - cell style, 6 - sheet properties, 7 - sheet clone", IsMandatory = true)]
        public int OutputType;
        [OSStructureField(DataType = OSDataType.Text, Description = "Name of the Sheet.", IsMandatory = true)]
        public string SheetName;
        [OSStructureField(Description = "ExcelTable record", IsMandatory = false)]
        public ExcelTable ExcelTable;
        [OSStructureField(Description = "ExcelCellValue record", IsMandatory = false)]
        public ExcelCellValue ExcelCellValue;
        [OSStructureField(Description = "ExcelPicture record", IsMandatory = false)]
        public ExcelPicture ExcelPicture;
        [OSStructureField(Description = "ExcelCellStyle record", IsMandatory = false)]
        public ExcelCellStyle ExcelCellStyle;
        [OSStructureField(Description = "ExcelMergeCells record", IsMandatory = false)]
        public ExcelMergeCells ExcelMergeCells;
        [OSStructureField(Description = "ExcelCloneSheet record", IsMandatory = false)]
        public ExcelCloneSheet ExcelCloneSheet;
        [OSStructureField(Description = "ExcelSheetProperties record", IsMandatory = false)]
        public ExcelSheetProperties ExcelSheetProperties;
    }

    [OSStructure(Description = "Table to insert in a specified position of a sheet.")]
    public struct ExcelTable
    {
        [OSStructureField(DataType = OSDataType.Integer, Description = "Row index", IsMandatory = true)]
        public int Row;
        [OSStructureField(DataType = OSDataType.Integer, Description = "Column index", IsMandatory = true)]
        public int Column;
        [OSStructureField(DataType = OSDataType.BinaryData, Description = "Binary Data to insert", IsMandatory = true)]
        public byte[] Binary;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "If True, discards the header of the table", IsMandatory = true)]
        public bool DiscardHeader;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "If True, uses the template cells data format", IsMandatory = true)]
        public bool UseTemplateCellsDataFormat;
    }

    [OSStructure(Description = "Value to set on a specified position (cell).")]
    public struct ExcelCellValue
    {
        [OSStructureField(DataType = OSDataType.Integer, Description = "Row index", IsMandatory = true)]
        public int Row;
        [OSStructureField(DataType = OSDataType.Integer, Description = "Column index", IsMandatory = true)]
        public int Column;
        [OSStructureField(DataType = OSDataType.Integer, Description = "1 - text, 2 - integer, 3 - decimal, 4 -boolean, 5 - datetime, 6 - date, 7 - time, 8 - formula", IsMandatory = true)]
        public int CellValueType;
        [OSStructureField(DataType = OSDataType.Text, Description = "Text value for the cell", IsMandatory = false)]
        public string TextValue;
        [OSStructureField(DataType = OSDataType.Integer, Description = "Integer value for the cell", IsMandatory = false)]
        public int IntegerValue;
        [OSStructureField(DataType = OSDataType.Decimal, Description = "Decimal value for the cell", IsMandatory = false)]
        public decimal DecimalValue;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "Boolean value for the cell", IsMandatory = false)]
        public bool BooleanValue;
        [OSStructureField(DataType = OSDataType.DateTime, Description = "Date Time value for the cell", IsMandatory = false)]
        public DateTime DateTimeValue;
        [OSStructureField(DataType = OSDataType.Date, Description = "Date value for the cell", IsMandatory = false)]
        public DateTime DateValue;
        [OSStructureField(DataType = OSDataType.Time, Description = "Time value for the cell", IsMandatory = false)]
        public DateTime TimeValue;
        [OSStructureField(DataType = OSDataType.Text, Description = "Formula for the cell", IsMandatory = false)]
        public string Formula;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "If True, uses the template cells data format", IsMandatory = true)]
        public bool UseTemplateCellsDataFormat;
    }

    [OSStructure(Description = "Picture to insert in excel document.")]
    public struct ExcelPicture
    {
        [OSStructureField(DataType = OSDataType.Integer, Description = "Row index", IsMandatory = true)]
        public int Row;
        [OSStructureField(DataType = OSDataType.Integer, Description = "Column index", IsMandatory = true)]
        public int Column;
        [OSStructureField(DataType = OSDataType.BinaryData, Description = "Picture to insert", IsMandatory = true)]
        public byte[] PictureBinary;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "If True, fits the picture within the cell of the specified coordinate. Otherwise, places the image across muliple cells while keeping its original size.", IsMandatory = false)]
        public bool FitWithinCell;
    }

    [OSStructure(Description = "Styles to apply to a specified position(cell).")]
    public struct ExcelCellStyle
    {
        [OSStructureField(DataType = OSDataType.Integer, Description = "Row index", IsMandatory = true)]
        public int Row;
        [OSStructureField(DataType = OSDataType.Integer, Description = "Column index", IsMandatory = true)]
        public int Column;
        [OSStructureField(DataType = OSDataType.Text, Description = "Cell color", IsMandatory = false)]
        public string CellColor;
        [OSStructureField(DataType = OSDataType.Text, Description = "Cell font color", IsMandatory = false)]
        public string CellFontColor;
        [OSStructureField(DataType = OSDataType.Text, Description = "Cell border bottom color", IsMandatory = false)]
        public string CellBorderBottomColor;
        [OSStructureField(DataType = OSDataType.Text, Description = "Cell border top color", IsMandatory = false)]
        public string CellBorderTopColor;
        [OSStructureField(DataType = OSDataType.Text, Description = "Cell border right color", IsMandatory = false)]
        public string CellBorderRightColor;
        [OSStructureField(DataType = OSDataType.Text, Description = "Cell border left color", IsMandatory = false)]
        public string CellBorderLeftColor;
        [OSStructureField(DataType = OSDataType.Integer, Description = "Defines the cell horizontal alignment:\r\n1 - Center\r\n2 - Left\r\n3 - Right", IsMandatory = false)]
        public int CellHorizontalAlignment;
        [OSStructureField(DataType = OSDataType.Integer, Description = "Defines the cell vertical alignment:\r\n1 - Bottom\r\n2 - Center\r\n3 - Top", IsMandatory = false)]
        public int CellVerticalAlignment;
    }

    [OSStructure(Description = "Merge specified cells in Excel file.")]
    public struct ExcelMergeCells
    {
        [OSStructureField(DataType = OSDataType.Integer, Description = "Row From (Row index)", IsMandatory = true)]
        public int Row;
        [OSStructureField(DataType = OSDataType.Integer, Description = "Column From (Column index)", IsMandatory = true)]
        public int Column;
        [OSStructureField(DataType = OSDataType.Integer, Description = "Row To (Row index)", IsMandatory = true)]
        public int RowTo;
        [OSStructureField(DataType = OSDataType.Integer, Description = "Column To (Column index)", IsMandatory = true)]
        public int ColumnTo;
    }

    [OSStructure(Description = "Cloning of a template sheet.")]
    public struct ExcelCloneSheet
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "Name of the Sheet from the Template which should be used as a Template for new one.", IsMandatory = true)]
        public string SheetToClone;
    }

    [OSStructure(Description = "Properties of an excel sheet.")]
    public struct ExcelSheetProperties
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "Password", IsMandatory = false)]
        public string Password;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "If True, unlocks formatting", IsMandatory = false)]
        public bool UnlockFormatting;
    }
}
