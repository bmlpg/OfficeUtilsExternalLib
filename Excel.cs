﻿using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.Extractor;
using NPOI.XSSF.UserModel;

namespace OfficeUtilsExternalLib
{
    internal class Excel
    {
        public enum ExcelOutputType : int
        {
            Table = 1,
            CellValue = 2,
            Picture = 3,
            MergeCells = 4,
            CellStyle = 5,
            SheetProperties = 6,
            SheetClone = 7
        }

        public enum CellValueType : int
        {
            Text = 1,
            Integer = 2,
            Decimal = 3,
            Boolean = 4,
            DateTime = 5,
            Date = 6,
            Time = 7,
            Formula = 8
        }


        public static byte[] GenerateExcelFile(ExcelStructures.ExcelFile excelFile)
        {
            XSSFWorkbook outputWorkbook;

            if (excelFile.Template.Length > 0)
            {
                MemoryStream templateStream = new MemoryStream(excelFile.Template);
                outputWorkbook = new XSSFWorkbook(templateStream);
            }
            else
            {
                outputWorkbook = new XSSFWorkbook();
            }

            ProcessExcelOutputs(excelFile, outputWorkbook);
            XSSFFormulaEvaluator.EvaluateAllFormulaCells(outputWorkbook);
            outputWorkbook.SetForceFormulaRecalculation(true);
            if (excelFile.LockStructure)
            {
                ExcelLockStructure(outputWorkbook);
            }
            MemoryStream outputStream = new MemoryStream();
            outputWorkbook.Write(outputStream);
            byte[] excelBinary = outputStream.ToArray();

            return excelBinary;

        } // MssCreateExcelFile

        private static void ProcessExcelOutputs(ExcelStructures.ExcelFile excelFile, XSSFWorkbook outputWorkbook)
        {
            List<string> sheetsToRemove = new List<string>();

            //process sheet cloning operations
            for (int i = 0; i < excelFile.ExcelOutputs.Count; i++)
            {
                ExcelStructures.ExcelOutput excelOutput = excelFile.ExcelOutputs[i];

                if (excelOutput.OutputType == (int)ExcelOutputType.SheetClone)
                {
                    string sheetToCloneName = excelOutput.ExcelCloneSheet.SheetToClone;

                    ISheet sheetToClone = outputWorkbook.GetSheet(sheetToCloneName);
                    outputWorkbook.CloneSheet(outputWorkbook.GetSheetIndex(sheetToClone), excelOutput.SheetName);

                    if (!sheetsToRemove.Contains(sheetToCloneName))
                    {
                        sheetsToRemove.Add(sheetToCloneName);
                    }
                }
            }


            //remove template (cloned) sheets
            foreach (string templateSheet in sheetsToRemove)
            {
                outputWorkbook.RemoveSheetAt(outputWorkbook.GetSheetIndex(outputWorkbook.GetSheet(templateSheet)));
            }

            //process all other operations
            for (int i = 0; i < excelFile.ExcelOutputs.Count; i++)
            {
                ExcelStructures.ExcelOutput excelOutput = excelFile.ExcelOutputs[i];
                ProcessExcelOutput(excelOutput, outputWorkbook);
            }

        }

        private static void ProcessExcelOutput(ExcelStructures.ExcelOutput excelOutput, XSSFWorkbook outputWorkbook)
        {
            ISheet outputSheet = outputWorkbook.GetSheet(excelOutput.SheetName);
            if (outputSheet == null)
            {
                outputSheet = outputWorkbook.CreateSheet(excelOutput.SheetName);
            }


            if (excelOutput.OutputType == (int)ExcelOutputType.Table)
            {
                ProcessExcelTable(excelOutput.ExcelTable, outputWorkbook, outputSheet);
            }
            else if (excelOutput.OutputType == (int)ExcelOutputType.CellValue)
            {
                ProcessExcelCellValue(excelOutput.ExcelCellValue, outputWorkbook, outputSheet);
            }
            else if (excelOutput.OutputType == (int)ExcelOutputType.Picture)
            {
                ProcessExcelPicture(excelOutput.ExcelPicture, outputWorkbook, outputSheet);
            }
            else if (excelOutput.OutputType == (int)ExcelOutputType.MergeCells)
            {
                ProcessExcelMergeCells(excelOutput.ExcelMergeCells, outputSheet);
            }
            else if (excelOutput.OutputType == (int)ExcelOutputType.CellStyle)
            {
                ProcessExcelCellStyle(excelOutput.ExcelCellStyle, outputWorkbook, outputSheet);
            }
            else if (excelOutput.OutputType == (int)ExcelOutputType.SheetProperties)
            {
                ProcessExcelSheetProperties(excelOutput.SheetName, excelOutput.ExcelSheetProperties, outputWorkbook);
            }
            else if (excelOutput.OutputType == (int)ExcelOutputType.SheetClone)
            {
                //ignore record
            }
            else
            {
                throw new Exception("Invalid Output Type: " + excelOutput.OutputType);
            }

        }

        private static void CloneFont(XSSFFont newFont, XSSFFont oldFont, bool withColor)
        {
            newFont.IsBold = oldFont.IsBold;
            if (oldFont.Charset != 0)
            {
                newFont.Charset = oldFont.Charset;
                newFont.SetCharSet(oldFont.Charset);
            }
            newFont.Family = oldFont.Family;
            newFont.FontHeight = oldFont.FontHeight;
            newFont.FontHeightInPoints = oldFont.FontHeightInPoints;
            newFont.FontName = oldFont.FontName;
            newFont.IsBold = oldFont.IsBold;
            newFont.IsItalic = oldFont.IsItalic;
            newFont.IsStrikeout = oldFont.IsStrikeout;
            newFont.TypeOffset = oldFont.TypeOffset;
            newFont.Underline = oldFont.Underline;
            newFont.SetScheme(oldFont.GetScheme());
            if (withColor)
            {
                newFont.Color = oldFont.Color;
                newFont.SetThemeColor(oldFont.GetThemeColor());
            }
        }



        private static void ProcessExcelTable(ExcelStructures.ExcelTable excelTable, XSSFWorkbook outputWorkbook, ISheet outputSheet)
        {
            MemoryStream inputStream = new MemoryStream(excelTable.Binary);

            XSSFWorkbook inputWorkbook = new XSSFWorkbook(inputStream);

            ISheet inputSheet = inputWorkbook.GetSheetAt(0);
            Dictionary<int, ICellStyle> cStyles = new Dictionary<int, ICellStyle>();
            int n = 0;
            if (excelTable.DiscardHeader == true)
            {
                n = 1;
            }

            for (int j = n; j <= inputSheet.LastRowNum; j++)
            {

                IRow inputRow = inputSheet.GetRow(j);
                IRow outputRow = outputSheet.GetRow(j - n + excelTable.Row);
                if (outputRow == null)
                {
                    outputRow = outputSheet.CreateRow(j - n + excelTable.Row);
                }
                for (int k = 0; k < inputRow.LastCellNum; k++)
                {
                    NPOI.SS.UserModel.ICell inputCell = inputRow.GetCell(k);

                    if (inputCell != null)
                    {
                        NPOI.SS.UserModel.ICell outputCell = outputRow.GetCell(k + excelTable.Column);
                        if (outputCell == null)
                        {
                            outputCell = outputRow.CreateCell(k + excelTable.Column);

                            outputCell.CellStyle = outputSheet.GetColumnStyle(k + excelTable.Column);

                            /*
                            XSSFSheet xssfSheet = (XSSFSheet)outputSheet;
                            ColumnHelper ch = xssfSheet.GetColumnHelper();
                            if (ch != null)
                            {
                                CT_Col col = ch.GetColumn(k + excelTable.Column, false);
                                if (col != null)
                                {

                                    if (col.IsSetStyle())
                                    {
                                        outputCell.CellStyle = outputSheet.GetColumnStyle(k + excelTable.Column);
                                    }
                                }
                            }
                            */

                            //create reusable style for the column 
                            if (!cStyles.TryGetValue(k, out ICellStyle style) && !excelTable.UseTemplateCellsDataFormat)
                            {
                                ICellStyle cellstyle = outputWorkbook.CreateCellStyle();
                                cellstyle.CloneStyleFrom(outputCell.CellStyle);

                                if (!excelTable.UseTemplateCellsDataFormat)
                                    cellstyle.DataFormat = outputWorkbook.CreateDataFormat().GetFormat(inputCell.CellStyle.GetDataFormatString());

                                cStyles.Add(k, cellstyle);
                            }

                            if (!excelTable.UseTemplateCellsDataFormat)
                            {
                                if (cStyles.TryGetValue(k, out ICellStyle tempStyle))
                                {
                                    outputCell.CellStyle = tempStyle;
                                }
                            }
                        }
                        else
                        {
                            if (!excelTable.UseTemplateCellsDataFormat)
                            {
                                outputCell.CellStyle.DataFormat = outputWorkbook.CreateDataFormat().GetFormat(inputCell.CellStyle.GetDataFormatString());
                            }
                        }

                        // Set the cell data value
                        switch (inputCell.CellType)
                        {
                            case CellType.Blank:
                                outputCell.SetCellValue(inputCell.StringCellValue);
                                break;
                            case CellType.Boolean:
                                outputCell.SetCellValue(inputCell.BooleanCellValue);
                                break;
                            case CellType.Error:
                                outputCell.SetCellErrorValue(inputCell.ErrorCellValue);
                                break;
                            case CellType.Formula:
                                outputCell.SetCellFormula(inputCell.CellFormula);
                                break;
                            case CellType.Numeric:
                                outputCell.SetCellValue(inputCell.NumericCellValue);
                                break;
                            case CellType.String:
                                outputCell.SetCellValue(inputCell.StringCellValue);
                                break;
                            case CellType.Unknown:
                                outputCell.SetCellValue(inputCell.StringCellValue);
                                break;
                        }

                    }

                }
            }

        }

        private static void ProcessExcelCellValue(ExcelStructures.ExcelCellValue excelCell, XSSFWorkbook outputWorkbook, ISheet outputSheet)
        {
            IRow outputRow = outputSheet.GetRow(excelCell.Row);
            if (outputRow == null)
            {
                outputRow = outputSheet.CreateRow(excelCell.Row);
            }

            NPOI.SS.UserModel.ICell outputCell = outputRow.GetCell(excelCell.Column);
            if (outputCell == null)
            {
                outputCell = outputRow.CreateCell(excelCell.Column);
                if (excelCell.UseTemplateCellsDataFormat)
                {
                    outputCell.CellStyle = outputSheet.GetColumnStyle(outputCell.ColumnIndex);

                    /*
                    XSSFSheet xssfSheet = (XSSFSheet)outputSheet;
                    ColumnHelper ch = xssfSheet.GetColumnHelper();
                    if (ch != null)
                    {
                        CT_Col col = ch.GetColumn(outputCell.ColumnIndex, false);
                        if (col != null)
                        {

                            if (col.IsSetStyle())
                            {
                                //ICellStyle cellstyle = outputWorkbook.CreateCellStyle();
                                //cellstyle.CloneStyleFrom(outputSheet.GetColumnStyle(k + ssExcelTable.ssSTExcelTable.ssColumn));
                                //outputCell.CellStyle = cellstyle;
                                outputCell.CellStyle = outputSheet.GetColumnStyle(outputCell.ColumnIndex);
                            }
                        }
                    }
                    */
                }
            }


            switch (excelCell.CellValueType)
            {
                case (int)CellValueType.Text:
                    if (!excelCell.UseTemplateCellsDataFormat)
                    {
                        ICellStyle cellStyle = outputWorkbook.CreateCellStyle();
                        cellStyle.DataFormat = (short)BuiltinFormats.GetBuiltinFormat("General");
                        outputCell.CellStyle = cellStyle;
                    }
                    outputCell.SetCellValue(excelCell.TextValue);
                    break;
                case (int)CellValueType.Integer:
                    if (!excelCell.UseTemplateCellsDataFormat)
                    {
                        ICellStyle cellStyle = outputWorkbook.CreateCellStyle();
                        cellStyle.DataFormat = (short)BuiltinFormats.GetBuiltinFormat("General");
                        outputCell.CellStyle = cellStyle;
                    }
                    outputCell.SetCellValue(excelCell.IntegerValue);
                    break;
                case (int)CellValueType.Decimal:
                    if (!excelCell.UseTemplateCellsDataFormat)
                    {
                        ICellStyle cellStyle = outputWorkbook.CreateCellStyle();
                        cellStyle.DataFormat = (short)BuiltinFormats.GetBuiltinFormat("General");
                        outputCell.CellStyle = cellStyle;
                    }
                    outputCell.SetCellValue(System.Convert.ToDouble(excelCell.DecimalValue));
                    break;
                case (int)CellValueType.Boolean:
                    if (!excelCell.UseTemplateCellsDataFormat)
                    {
                        ICellStyle cellStyle = outputWorkbook.CreateCellStyle();
                        cellStyle.DataFormat = (short)BuiltinFormats.GetBuiltinFormat("General");
                        outputCell.CellStyle = cellStyle;
                    }
                    outputCell.SetCellValue(excelCell.BooleanValue);
                    break;
                case (int)CellValueType.DateTime:
                    if (!excelCell.UseTemplateCellsDataFormat)
                    {
                        ICellStyle cellStyle = outputWorkbook.CreateCellStyle();
                        cellStyle.DataFormat = outputWorkbook.CreateDataFormat().GetFormat("yyyy-MM-dd HH:mm:ss");
                        outputCell.CellStyle = cellStyle;
                    }
                    outputCell.SetCellValue(excelCell.DateTimeValue);
                    break;
                case (int)CellValueType.Date:
                    if (!excelCell.UseTemplateCellsDataFormat)
                    {
                        ICellStyle cellStyle = outputWorkbook.CreateCellStyle();
                        cellStyle.DataFormat = outputWorkbook.CreateDataFormat().GetFormat("yyyy-MM-dd");
                        outputCell.CellStyle = cellStyle;
                    }
                    outputCell.SetCellValue(excelCell.DateValue);
                    break;
                case (int)CellValueType.Time:
                    if (!excelCell.UseTemplateCellsDataFormat)
                    {
                        ICellStyle cellStyle = outputWorkbook.CreateCellStyle();
                        cellStyle.DataFormat = outputWorkbook.CreateDataFormat().GetFormat("HH:mm:ss");
                        outputCell.CellStyle = cellStyle;
                    }
                    outputCell.SetCellValue(excelCell.TimeValue);
                    break;
                case (int)CellValueType.Formula:
                    outputCell.SetCellFormula(excelCell.Formula);
                    break;
                default:
                    throw new Exception("Invalid Value Type: " + excelCell.CellValueType);
            }

        }

        private static void ProcessExcelPicture(ExcelStructures.ExcelPicture excelPicture, XSSFWorkbook outputWorkbook, ISheet outputSheet)
        {
            PictureType pictureType = Utils.GetPictureType(excelPicture.PictureBinary);

            int pictureIdx = outputWorkbook.AddPicture(excelPicture.PictureBinary, pictureType);

            IDrawing patriarch = outputSheet.CreateDrawingPatriarch();

            ICreationHelper helper = outputWorkbook.GetCreationHelper();

            IClientAnchor anchor = helper.CreateClientAnchor();
            anchor.Col1 = excelPicture.Column;
            anchor.Row1 = excelPicture.Row;

            if (excelPicture.FitWithinCell)
            {
                anchor.Col2 = anchor.Col1 + 1;
                anchor.Row2 = anchor.Row1 + 1;
            }

            XSSFPicture picture = (XSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);

            if (!excelPicture.FitWithinCell)
            {
                picture.Resize();
            }
        }

        private static void ProcessExcelCellStyle(ExcelStructures.ExcelCellStyle excelCellStyle, XSSFWorkbook outputWorkbook, ISheet outputSheet)
        {
            IRow outputRow = outputSheet.GetRow(excelCellStyle.Row);
            if (outputRow == null)
            {
                outputRow = outputSheet.CreateRow(excelCellStyle.Row);
            }

            NPOI.SS.UserModel.ICell outputCell = outputRow.GetCell(excelCellStyle.Column);
            if (outputCell == null)
            {
                outputCell = outputRow.CreateCell(excelCellStyle.Column);
            }

            XSSFCellStyle cellStyle = (XSSFCellStyle)outputWorkbook.CreateCellStyle();
            cellStyle.CloneStyleFrom(outputCell.CellStyle);

            if (excelCellStyle.CellColor != "")
            {
                cellStyle.SetFillForegroundColor(new XSSFColor(SixLabors.ImageSharp.Color.ParseHex(excelCellStyle.CellColor)));
                //cellStyle.FillForegroundColor = IndexedColors.Orange.Index; 
                cellStyle.FillPattern = FillPattern.SolidForeground;
            }

            if (excelCellStyle.CellFontColor != "")
            {
                XSSFFont NEWfont = (XSSFFont)outputWorkbook.CreateFont();
                CloneFont(NEWfont, cellStyle.GetFont(), false);
                NEWfont.SetColor(new XSSFColor(SixLabors.ImageSharp.Color.ParseHex(excelCellStyle.CellFontColor)));
                cellStyle.SetFont(NEWfont);
            }

            if (excelCellStyle.CellBorderBottomColor != "" || excelCellStyle.CellBorderLeftColor != ""
                || excelCellStyle.CellBorderRightColor != "" || excelCellStyle.CellBorderTopColor != "")
            {
                if (outputCell.IsMergedCell)
                {
                    int startRow = excelCellStyle.Row;
                    int endRow = 0;
                    int startColumn = excelCellStyle.Column;
                    int endColumn = 0;
                    CellRangeAddress cellRange;
                    for (int a = 0; a < outputSheet.NumMergedRegions; a++)
                    {
                        cellRange = outputSheet.GetMergedRegion(a);
                        if (startRow == cellRange.FirstRow && startColumn == cellRange.FirstColumn)
                        {
                            endRow = cellRange.LastRow;
                            endColumn = cellRange.LastColumn;
                        }
                    }
                    for (int i = startRow; i <= endRow; i++)
                    {
                        for (int j = startColumn; j <= endColumn; j++)
                        {
                            if (i == startRow && j == startColumn)
                            { }
                            else
                            {
                                IRow row = outputSheet.GetRow(i) ?? outputSheet.CreateRow(i);
                                NPOI.SS.UserModel.ICell mergedOutputCell = row.GetCell(j) ?? row.CreateCell(j);
                                XSSFCellStyle cellStyleMerged = (XSSFCellStyle)outputWorkbook.CreateCellStyle();
                                cellStyleMerged.CloneStyleFrom(mergedOutputCell.CellStyle);
                                if (excelCellStyle.CellBorderBottomColor != "")
                                {
                                    cellStyleMerged.BorderBottom = BorderStyle.Medium;
                                    cellStyleMerged.SetBottomBorderColor(new XSSFColor(SixLabors.ImageSharp.Color.ParseHex(excelCellStyle.CellBorderBottomColor)));
                                }
                                if (excelCellStyle.CellBorderTopColor != "")
                                {
                                    cellStyleMerged.BorderTop = BorderStyle.Medium;
                                    cellStyleMerged.SetTopBorderColor(new XSSFColor(SixLabors.ImageSharp.Color.ParseHex(excelCellStyle.CellBorderTopColor)));
                                }
                                if (excelCellStyle.CellBorderLeftColor != "")
                                {
                                    cellStyleMerged.BorderLeft = BorderStyle.Medium;
                                    cellStyleMerged.SetLeftBorderColor(new XSSFColor(SixLabors.ImageSharp.Color.ParseHex(excelCellStyle.CellBorderLeftColor)));
                                }
                                if (excelCellStyle.CellBorderRightColor != "")
                                {
                                    cellStyleMerged.BorderRight = BorderStyle.Medium;
                                    cellStyleMerged.SetRightBorderColor(new XSSFColor(SixLabors.ImageSharp.Color.ParseHex(excelCellStyle.CellBorderRightColor)));
                                }
                                mergedOutputCell.CellStyle = cellStyleMerged;
                            }
                        }
                    }

                }

                if (excelCellStyle.CellBorderBottomColor != "")
                {
                    cellStyle.BorderBottom = BorderStyle.Medium;
                    cellStyle.SetBottomBorderColor(new XSSFColor(SixLabors.ImageSharp.Color.ParseHex(excelCellStyle.CellBorderBottomColor)));
                }
                if (excelCellStyle.CellBorderTopColor != "")
                {
                    cellStyle.BorderTop = BorderStyle.Medium;
                    cellStyle.SetTopBorderColor(new XSSFColor(SixLabors.ImageSharp.Color.ParseHex(excelCellStyle.CellBorderTopColor)));
                }
                if (excelCellStyle.CellBorderLeftColor != "")
                {
                    cellStyle.BorderLeft = BorderStyle.Medium;
                    cellStyle.SetLeftBorderColor(new XSSFColor(SixLabors.ImageSharp.Color.ParseHex(excelCellStyle.CellBorderLeftColor)));
                }
                if (excelCellStyle.CellBorderRightColor != "")
                {
                    cellStyle.BorderRight = BorderStyle.Medium;
                    cellStyle.SetRightBorderColor(new XSSFColor(SixLabors.ImageSharp.Color.ParseHex(excelCellStyle.CellBorderRightColor)));
                }

            }

            if (excelCellStyle.CellHorizontalAlignment > 0)
            {
                if (excelCellStyle.CellHorizontalAlignment == 1)
                {
                    cellStyle.Alignment = HorizontalAlignment.Center;
                }
                else if (excelCellStyle.CellHorizontalAlignment == 2)
                {
                    cellStyle.Alignment = HorizontalAlignment.Left;
                }
                else if (excelCellStyle.CellHorizontalAlignment == 3)
                {
                    cellStyle.Alignment = HorizontalAlignment.Right;
                }
            }

            if (excelCellStyle.CellVerticalAlignment > 0)
            {
                if (excelCellStyle.CellVerticalAlignment == 1)
                {
                    cellStyle.VerticalAlignment = VerticalAlignment.Bottom;
                }
                else if (excelCellStyle.CellVerticalAlignment == 2)
                {
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                }
                else if (excelCellStyle.CellVerticalAlignment == 3)
                {
                    cellStyle.VerticalAlignment = VerticalAlignment.Top;
                }

            }

            outputCell.CellStyle = cellStyle;
        }

        private static void ProcessExcelMergeCells(ExcelStructures.ExcelMergeCells excelMergeCells, ISheet outputSheet)
        {


            outputSheet.AddMergedRegion(
                new CellRangeAddress(
                    excelMergeCells.Row,
                    excelMergeCells.RowTo,
                    excelMergeCells.Column,
                    excelMergeCells.ColumnTo
                )
            );
        }

        private static void ProcessExcelSheetProperties(string sheetName, ExcelStructures.ExcelSheetProperties sheetProperties, XSSFWorkbook outputWorkbook)
        {
            ISheet outputSheet = outputWorkbook.GetSheet(sheetName);

            if (outputSheet != null)

            {
                /* Set Password */

                outputSheet.ProtectSheet(sheetProperties.Password);
                XSSFSheet sheet = ((XSSFSheet)outputSheet);

                if (sheetProperties.UnlockFormatting) /* Unlock formating */
                {
                    sheet.LockFormatColumns(false);
                    sheet.LockFormatCells(false);
                    sheet.LockFormatRows(false);
                    sheet.LockObjects(false);
                    sheet.LockScenarios(false);
                }

            }

        }

        private static void ExcelLockStructure(XSSFWorkbook outputWorkbook)
        {
            outputWorkbook.LockStructure();
        }

        public static string ExtractExcelFileContent(byte[] ssExcelBinary)
        {
            MemoryStream memoryStream = new MemoryStream(ssExcelBinary);
            XSSFWorkbook workbook = new XSSFWorkbook(memoryStream);
            XSSFExcelExtractor extractor = new XSSFExcelExtractor(workbook);
            return extractor.Text;
        }
    }
}
