using OutSystems.ExternalLibraries.SDK;

namespace OfficeUtilsExternalLib.WordStructures
{
    [OSStructure(Description = "Contains template document, and a list of all replacements to be performed.")]
    public struct WordFile
    {
        [OSStructureField(Description = "WordOutput record list", IsMandatory = true)]
        public List<WordOutput> WordOutputs;
        [OSStructureField(DataType = OSDataType.BinaryData, Description = "Template word file", IsMandatory = true)]
        public byte[] Template;
    }

    [OSStructure(Description = "Replacements to be performed in word document.")]
    public struct WordOutput
    {
        [OSStructureField(DataType = OSDataType.Integer, Description = "1 - text, 2 - legacy table, 3 - picture, 4 -  table ", IsMandatory = true)]
        public int OutputType;
        [OSStructureField(DataType = OSDataType.Text, Description = "Placeholder to replace", IsMandatory = true)]
        public string Placeholder;
        [OSStructureField(Description = "Text to replace the placeholder with", IsMandatory = false)]
        public WordText WordText;
        [OSStructureField(Description = "WordLegacyTable record", IsMandatory = false)]
        public WordLegacyTable WordLegacyTable;
        [OSStructureField(Description = "WordTable record", IsMandatory = false)]
        public WordTable WordTable;
        [OSStructureField(Description = "Picture to replace the placeholder with", IsMandatory = false)]
        public WordPicture WordPicture;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "If True, deletes the placeholder", IsMandatory = false)]
        public bool DeletePlaceholder;
    }

    [OSStructure(Description = "Textual value to insert in the word document.")]
    public struct WordText
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "Text value to insert in the word document", IsMandatory = true)]
        public string Text;
        [OSStructureField(DataType = OSDataType.Text, Description = "Hyperlink to insert in the word document", IsMandatory = false)]
        public string Hyperlink;
    }

    [OSStructure(Description = "Table to insert in the word document.")]
    public struct WordLegacyTable
    {
        [OSStructureField(Description = "WordLegacyTableRow record list", IsMandatory = true)]
        public List<WordLegacyTableRow> TableRows;
    }

    [OSStructure(Description = "Row of a table to insert in word document.")]
    public struct WordLegacyTableRow
    {
        [OSStructureField(Description = "WordLegacyTableCell record list", IsMandatory = true)]
        public List<WordLegacyTableCell> TableCells;
    }

    [OSStructure(Description = "Represents the value to be placed in a table cell.")]
    public struct WordLegacyTableCell
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "Text value to insert in the word document", IsMandatory = true)]
        public string Value;
    }

    [OSStructure(Description = "Contains all rows to insert in an table of word document.")]
    public struct WordTable
    {
        [OSStructureField(Description = "WordTableRow record list", IsMandatory = true)]
        public List<WordTableRow> TableRows;
        [OSStructureField(DataType = OSDataType.Integer, Description = "Starts with 0. If you have header, the first row should be 1.", IsMandatory = true)]
        public int StartRow;
    }

    [OSStructure(Description = "Content to be be added to a table row.")]
    public struct WordTableRow
    {
        [OSStructureField(Description = "WordTableReplacement record list", IsMandatory = true)]
        public List<WordTableRowReplacement> RowReplacements;
    }

    [OSStructure(Description = "Rows to insert in the list of rows of a table to include in the word document.")]
    public struct WordTableRowReplacement
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "Placeholder to replace", IsMandatory = true)]
        public string Placeholder;
        [OSStructureField(DataType = OSDataType.BinaryData, Description = "Picture to insert in the placeholder", IsMandatory = false)]
        public byte[] Picture;
        [OSStructureField(DataType = OSDataType.Integer, Description = "Picture width in pixels", IsMandatory = false)]
        public int PictureWidth;
        [OSStructureField(DataType = OSDataType.Text, Description = "Text to replace the placeholder with", IsMandatory = false)]
        public string Text;
    }

    [OSStructure(Description = "Picture to insert in word document.")]
    public struct WordPicture
    {
        [OSStructureField(DataType = OSDataType.BinaryData, Description = "Picture to insert", IsMandatory = true)]
        public byte[] Picture;
        [OSStructureField(DataType = OSDataType.Integer, Description = "Width in pixels", IsMandatory = false)]
        public int Width;
    }
}
