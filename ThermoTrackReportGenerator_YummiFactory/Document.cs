using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

using _ooxml_ = DocumentFormat.OpenXml;
using _drawing_ = DocumentFormat.OpenXml.Drawing;
using _drawing_charts_ = DocumentFormat.OpenXml.Drawing.Charts;
using _drawing_charts_2010_ = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using _drawing_pictures_ = DocumentFormat.OpenXml.Drawing.Pictures;
using _drawing_word_ = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using _package_ = DocumentFormat.OpenXml.Packaging;
using _valid_ = DocumentFormat.OpenXml.Validation;
using _word_ = DocumentFormat.OpenXml.Wordprocessing;

namespace ThermoTrackReportGenerator_YummiFactory
{
    public class Document
    {
        private class DocumentSection
        {
            public _word_.SectionProperties SectionProperties { get; set; }

            public List<_ooxml_.OpenXmlCompositeElement> Elements { get; set; }
            public int ElementsIndexCurrent { get; set; }

            public DocumentSection()
            {
                SectionProperties = new _word_.SectionProperties();

                Elements = new List<_ooxml_.OpenXmlCompositeElement>();

                ElementsIndexCurrent = 0;
            }
        }

        public enum ListFormatType
        {
            Bullet,
            Decimal,
            LowerRoman
        }

        public enum AlignmentHorizontal
        {
            Left,
            Centre,
            Right
        }

        public enum UnitType
        {
            None,
            Temperature_Celsius
        }

        private MemoryStream _documentStream;
        private _package_.WordprocessingDocument _document;

        private _package_.DocumentSettingsPart _documentSettingsPart;
        private _package_.MainDocumentPart _mainDocumentPart;

        private _word_.Body _body;

        private List<DocumentSection> _sections;
        private int _sectionsIndexCurrent = 0;

        private _package_.NumberingDefinitionsPart _numberingDefinitionsPart;
        private int _idListFormatNext = 1;

        private uint _idDocElementNext = 1;
        private uint _idPictureNext = 1;

        private uint _idRelationshipNext = 1;

        public Document()
        {
            _documentStream = new MemoryStream();

            _document = _package_.WordprocessingDocument.Create(_documentStream, _ooxml_.WordprocessingDocumentType.Document);

            _mainDocumentPart = _document.AddMainDocumentPart();
            _mainDocumentPart.Document = new _word_.Document();

            _body = _mainDocumentPart.Document.AppendChild(new _word_.Body());

            _sections = new List<DocumentSection>();

            // Add the DocumentSettingsPart.

            string idRelationship = "rId" + _idRelationshipNext++;

            _documentSettingsPart = _mainDocumentPart.AddNewPart<_package_.DocumentSettingsPart>(idRelationship);
            _documentSettingsPart.Settings = new _word_.Settings();

            _documentSettingsPart.Settings.AppendChild(new _word_.UpdateFieldsOnOpen() { Val = true });

            // Add the NumberingDefinitionsPart.

            idRelationship = "rId" + _idRelationshipNext++;

            _numberingDefinitionsPart = _mainDocumentPart.AddNewPart<_package_.NumberingDefinitionsPart>(idRelationship);
            _numberingDefinitionsPart.Numbering = new _word_.Numbering();
        }

        ~Document()
        {
            _document.Dispose();
            _documentStream.Dispose();
        }

        public int AddSection()
        {
            _sections.Add(new DocumentSection());

            return _sectionsIndexCurrent++;
        }

        public int AddParagraph(int indexSection, AlignmentHorizontal alignment = AlignmentHorizontal.Left)
        {
            _word_.JustificationValues justificationValue = _word_.JustificationValues.Left;

            if (alignment == AlignmentHorizontal.Centre)
                justificationValue = _word_.JustificationValues.Center;

            else if (alignment == AlignmentHorizontal.Right)
                justificationValue = _word_.JustificationValues.Right;

            _word_.Paragraph paragraph = new _word_.Paragraph();
            _word_.ParagraphProperties paragraphProperties = paragraph.AppendChild(new _word_.ParagraphProperties());
            paragraphProperties.AppendChild(new _word_.SpacingBetweenLines()
            {
                After = "0",
                Line = "0",
                LineRule = _word_.LineSpacingRuleValues.AtLeast
            });
            paragraphProperties.AppendChild(new _word_.Justification() { Val = justificationValue });

            // Add element to section.

            _sections[indexSection].Elements.Add(paragraph);

            return _sections[indexSection].ElementsIndexCurrent++;
        }

        public void AddRun(int indexSection, int indexElement, bool isBreak = true, uint fontSize = 24)
        {
            _word_.Run run = _sections[indexSection].Elements[indexElement].AppendChild(new _word_.Run());
            _word_.RunProperties runProperties = run.AppendChild(new _word_.RunProperties());
            runProperties.AppendChild(new _word_.RunFonts() { Ascii = "Liberation Sans" });
            runProperties.AppendChild(new _word_.FontSize() { Val = Convert.ToString(fontSize) });

            if (isBreak)
                run.AppendChild(new _word_.Break());

            else
                run.AppendChild(new _word_.TabChar());
        }

        public void AddRun(int indexSection, int indexElement, string text, bool isUnderlined = false, bool isBold = false, uint fontSize = 24)
        {
            _word_.Run run = _sections[indexSection].Elements[indexElement].AppendChild(new _word_.Run());
            _word_.RunProperties runProperties = run.AppendChild(new _word_.RunProperties());
            runProperties.AppendChild(new _word_.RunFonts() { Ascii = "Liberation Sans" });

            if (isBold)
                runProperties.AppendChild(new _word_.Bold());

            runProperties.AppendChild(new _word_.FontSize() { Val = Convert.ToString(fontSize) });

            if (isUnderlined)
                runProperties.AppendChild(new _word_.Underline() { Val = _word_.UnderlineValues.Single });

            run.AppendChild(new _word_.Text()
            {
                Text = text,
                Space = _ooxml_.SpaceProcessingModeValues.Preserve
            });
        }

        public int AddToC(int indexSection, uint levels)
        {
            _word_.Paragraph paragraph = new _word_.Paragraph();

            {
                _word_.Run run = paragraph.AppendChild(new _word_.Run());
                run.AppendChild(new _word_.FieldChar() { FieldCharType = _word_.FieldCharValues.Begin });
            }

            {
                _word_.Run run = paragraph.AppendChild(new _word_.Run());
                run.AppendChild(new _word_.FieldCode()
                {
                    Space = _ooxml_.SpaceProcessingModeValues.Preserve,
                    Text = "TOC \\o \"1-" + levels + "\" \\f TOC"
                });
            }

            {
                _word_.Run run = paragraph.AppendChild(new _word_.Run());
                run.AppendChild(new _word_.FieldChar() { FieldCharType = _word_.FieldCharValues.Separate });
            }

            {
                _word_.Run run = paragraph.AppendChild(new _word_.Run());
                run.AppendChild(new _word_.FieldChar() { FieldCharType = _word_.FieldCharValues.End });
            }

            _sections[indexSection].Elements.Add(paragraph);

            return _sections[indexSection].ElementsIndexCurrent++;
        }

        public void AddToCTag(int indexSection, int indexElement, uint level, string text)
        {
            {
                _word_.Run run = _sections[indexSection].Elements[indexElement].AppendChild(new _word_.Run());
                run.AppendChild(new _word_.FieldChar() { FieldCharType = _word_.FieldCharValues.Begin });
            }

            {
                _word_.Run run = _sections[indexSection].Elements[indexElement].AppendChild(new _word_.Run());
                run.AppendChild(new _word_.FieldCode()
                {
                    Space = _ooxml_.SpaceProcessingModeValues.Preserve,
                    Text = "TC \"" + text + "\" \\f TOC \\l " + level
                });
            }

            {
                _word_.Run run = _sections[indexSection].Elements[indexElement].AppendChild(new _word_.Run());
                run.AppendChild(new _word_.FieldChar() { FieldCharType = _word_.FieldCharValues.Separate });
            }

            {
                _word_.Run run = _sections[indexSection].Elements[indexElement].AppendChild(new _word_.Run());
                run.AppendChild(new _word_.FieldChar() { FieldCharType = _word_.FieldCharValues.End });
            }
        }

        public int AddListFormat(ListFormatType listType = ListFormatType.Bullet, string suffix = ".", bool isBold = false, int levelCount = 1)
        {
            _word_.NumberFormatValues numberFormatValue = _word_.NumberFormatValues.Bullet;

            if (listType == ListFormatType.Decimal)
                numberFormatValue = _word_.NumberFormatValues.Decimal;

            if (listType == ListFormatType.LowerRoman)
                numberFormatValue = _word_.NumberFormatValues.LowerRoman;

            _word_.AbstractNum abstractNum = _numberingDefinitionsPart.Numbering.PrependChild(new _word_.AbstractNum() { AbstractNumberId = (_idListFormatNext - 1) });

            abstractNum.AppendChild(new _word_.MultiLevelType() { Val = _word_.MultiLevelValues.Multilevel });

            for (int i = 0; i < levelCount; ++i)
            {
                _word_.Level level = abstractNum.AppendChild(new _word_.Level() { LevelIndex = i });
                level.AppendChild(new _word_.StartNumberingValue() { Val = 1 });
                level.AppendChild(new _word_.NumberingFormat() { Val = numberFormatValue });

                if (numberFormatValue == _word_.NumberFormatValues.Bullet)
                    level.AppendChild(new _word_.LevelText() { Val = "\u2022" });
                else
                    level.AppendChild(new _word_.LevelText() { Val = "%" + (i + 1) + suffix });

                level.AppendChild(new _word_.LevelJustification() { Val = _word_.LevelJustificationValues.Left });

                _word_.PreviousParagraphProperties previousParagraphProperties = level.AppendChild(new _word_.PreviousParagraphProperties());
                previousParagraphProperties.AppendChild(new _word_.SpacingBetweenLines()
                {
                    After = "0",
                    Line = "0",
                    LineRule = _word_.LineSpacingRuleValues.AtLeast
                });
                previousParagraphProperties.AppendChild(new _word_.Indentation()
                {
                    Left = "720",
                    Hanging = "360"
                });

                _word_.NumberingSymbolRunProperties numberingSymbolRunProperties = level.AppendChild(new _word_.NumberingSymbolRunProperties());
                numberingSymbolRunProperties.AppendChild(new _word_.RunFonts()
                {
                    Ascii = "Liberation Sans",
                    ComplexScript = "Liberation Sans"
                });

                if (isBold)
                    numberingSymbolRunProperties.AppendChild(new _word_.Bold());
            }

            _word_.NumberingInstance numberingInstance = _numberingDefinitionsPart.Numbering.AppendChild(new _word_.NumberingInstance() { NumberID = _idListFormatNext });
            numberingInstance.AppendChild(new _word_.AbstractNumId() { Val = (_idListFormatNext - 1) });

            return _idListFormatNext++;
        }

        public int AddListItem(int indexSection, int idListFormat, int levelIndex = 0)
        {
            _word_.Paragraph paragraph = new _word_.Paragraph();

            _word_.ParagraphProperties paragraphProperties = paragraph.AppendChild(new _word_.ParagraphProperties());
            _word_.NumberingProperties numberingProperties = paragraphProperties.AppendChild(new _word_.NumberingProperties());
            numberingProperties.AppendChild(new _word_.NumberingLevelReference() { Val = levelIndex });
            numberingProperties.AppendChild(new _word_.NumberingId() { Val = idListFormat });

            // Add element to section.

            _sections[indexSection].Elements.Add(paragraph);

            return _sections[indexSection].ElementsIndexCurrent++;
        }

        public void AddHeader_Empty(int indexSection, bool isFirstPage = false)
        {
            _word_.HeaderFooterValues headerFooterValue = _word_.HeaderFooterValues.Default;

            if (isFirstPage)
                headerFooterValue = _word_.HeaderFooterValues.First;

            string idRelationship = "rId" + _idRelationshipNext++;

            _word_.HeaderReference headerReference = _sections[indexSection].SectionProperties.AppendChild(new _word_.HeaderReference()
            {
                Type = headerFooterValue,
                Id = idRelationship
            });

            _package_.HeaderPart headerPart = _mainDocumentPart.AddNewPart<_package_.HeaderPart>(idRelationship);
            headerPart.Header = new _word_.Header();
        }

        public void AddHeader(int indexSection, bool isFirstPage = false, AlignmentHorizontal alignment = AlignmentHorizontal.Left, bool hasPageNum = false, string text = "", bool hasSectionNum = false, int sectionNum = 1)
        {
            _word_.HeaderFooterValues headerFooterValue = _word_.HeaderFooterValues.Default;

            if (isFirstPage)
                headerFooterValue = _word_.HeaderFooterValues.First;

            _word_.JustificationValues justificationValue = _word_.JustificationValues.Left;

            if (alignment == AlignmentHorizontal.Centre)
                justificationValue = _word_.JustificationValues.Center;

            else if (alignment == AlignmentHorizontal.Right)
                justificationValue = _word_.JustificationValues.Right;

            string idRelationship = "rId" + _idRelationshipNext++;

            _word_.HeaderReference headerReference = _sections[indexSection].SectionProperties.AppendChild(new _word_.HeaderReference()
            {
                Type = headerFooterValue,
                Id = idRelationship
            });

            _package_.HeaderPart headerPart = _mainDocumentPart.AddNewPart<_package_.HeaderPart>(idRelationship);
            headerPart.Header = new _word_.Header();

            _word_.Table table = headerPart.Header.AppendChild(new _word_.Table());

            _word_.TableProperties tableProperties = table.AppendChild(new _word_.TableProperties());

            tableProperties.AppendChild(new _word_.TableWidth()
            {
                Width = "5973",
                Type = _word_.TableWidthUnitValues.Pct
            });

            tableProperties.AppendChild(new _word_.TableIndentation()
            {
                Width = -878,
                Type = _word_.TableWidthUnitValues.Dxa
            });

            _word_.TableBorders tableBorders = tableProperties.AppendChild(new _word_.TableBorders());
            tableBorders.AppendChild(new _word_.BottomBorder()
            {
                Val = _word_.BorderValues.Single,
                Size = 18,
                Space = 0,
                Color = "808080",
                ThemeColor = _word_.ThemeColorValues.Background1,
                ThemeShade = "80"
            });

            _word_.TableCellMarginDefault tableCellMarginDefault = tableProperties.AppendChild(new _word_.TableCellMarginDefault());
            tableCellMarginDefault.AppendChild(new _word_.TopMargin()
            {
                Width = "72",
                Type = _word_.TableWidthUnitValues.Dxa
            });
            tableCellMarginDefault.AppendChild(new _word_.TableCellLeftMargin()
            {
                Width = 115,
                Type = _word_.TableWidthValues.Dxa
            });
            tableCellMarginDefault.AppendChild(new _word_.BottomMargin()
            {
                Width = "72",
                Type = _word_.TableWidthUnitValues.Dxa
            });
            tableCellMarginDefault.AppendChild(new _word_.TableCellRightMargin()
            {
                Width = 115,
                Type = _word_.TableWidthValues.Dxa
            });

            tableProperties.AppendChild(new _word_.TableLook()
            {
                Val = "04A0",
                FirstRow = true,
                LastRow = false,
                FirstColumn = true,
                LastColumn = false,
                NoHorizontalBand = false,
                NoVerticalBand = true
            });

            _word_.TableGrid tableGrid = table.AppendChild(new _word_.TableGrid());
            tableGrid.AppendChild(new _word_.GridColumn() { Width = "11057" });

            _word_.TableRow tableRow = table.AppendChild(new _word_.TableRow());

            _word_.TableRowProperties tableRowProperties = tableRow.AppendChild(new _word_.TableRowProperties());
            tableRowProperties.AppendChild(new _word_.TableRowHeight() { Val = 64 });

            _word_.TableCell tableCell = tableRow.AppendChild(new _word_.TableCell());

            _word_.TableCellProperties tableCellProperties = tableCell.AppendChild(new _word_.TableCellProperties());
            tableCellProperties.AppendChild(new _word_.TableCellWidth()
            {
                Width = "11058",
                Type = _word_.TableWidthUnitValues.Dxa
            });

            _word_.Paragraph paragraph = tableCell.AppendChild(new _word_.Paragraph());

            _word_.ParagraphProperties paragraphProperties = paragraph.AppendChild(new _word_.ParagraphProperties());
            paragraphProperties.AppendChild(new _word_.SpacingBetweenLines()
            {
                After = "0",
                Line = "0",
                LineRule = _word_.LineSpacingRuleValues.AtLeast
            });
            paragraphProperties.AppendChild(new _word_.Justification() { Val = justificationValue });

            _word_.Run run = paragraph.AppendChild(new _word_.Run());

            _word_.RunProperties runProperties = run.AppendChild(new _word_.RunProperties());
            runProperties.AppendChild(new _word_.RunFonts() { Ascii = "Liberation Sans" });
            runProperties.AppendChild(new _word_.Bold());
            runProperties.AppendChild(new _word_.FontSize() { Val = "20" });

            if (hasPageNum)
            {
                run.AppendChild(new _word_.PageNumber());
                run.AppendChild(new _word_.TabChar());
            }

            else if (hasSectionNum)
            {
                run.AppendChild(new _word_.Text() { Text = Convert.ToString(sectionNum) });
                run.AppendChild(new _word_.TabChar());
            }

            if (text.Length > 0)
                run.AppendChild(new _word_.Text()
                {
                    Text = text,
                    Space = _ooxml_.SpaceProcessingModeValues.Preserve
                });
        }

        public void AddFooter(int indexSection, bool isFirstPage = false, AlignmentHorizontal alignment = AlignmentHorizontal.Left, bool hasPageNum = false, string text = "", bool hasSectionNum = false, int sectionNum = 1)
        {
            _word_.HeaderFooterValues headerFooterValue = _word_.HeaderFooterValues.Default;

            if (isFirstPage)
                headerFooterValue = _word_.HeaderFooterValues.First;

            _word_.JustificationValues justificationValue = _word_.JustificationValues.Left;

            if (alignment == AlignmentHorizontal.Centre)
                justificationValue = _word_.JustificationValues.Center;

            else if (alignment == AlignmentHorizontal.Right)
                justificationValue = _word_.JustificationValues.Right;

            string idRelationship = "rId" + _idRelationshipNext++;

            _word_.FooterReference footerReference = _sections[indexSection].SectionProperties.AppendChild(new _word_.FooterReference()
            {
                Type = headerFooterValue,
                Id = idRelationship
            });

            _package_.FooterPart footerPart = _mainDocumentPart.AddNewPart<_package_.FooterPart>(idRelationship);
            footerPart.Footer = new _word_.Footer();

            _word_.Table table = footerPart.Footer.AppendChild(new _word_.Table());

            _word_.TableProperties tableProperties = table.AppendChild(new _word_.TableProperties());

            tableProperties.AppendChild(new _word_.TableWidth()
            {
                Width = "5973",
                Type = _word_.TableWidthUnitValues.Pct
            });

            tableProperties.AppendChild(new _word_.TableIndentation()
            {
                Width = -878,
                Type = _word_.TableWidthUnitValues.Dxa
            });

            _word_.TableBorders tableBorders = tableProperties.AppendChild(new _word_.TableBorders());
            tableBorders.AppendChild(new _word_.TopBorder()
            {
                Val = _word_.BorderValues.Single,
                Size = 18,
                Space = 0,
                Color = "808080",
                ThemeColor = _word_.ThemeColorValues.Background1,
                ThemeShade = "80"
            });

            _word_.TableCellMarginDefault tableCellMarginDefault = tableProperties.AppendChild(new _word_.TableCellMarginDefault());
            tableCellMarginDefault.AppendChild(new _word_.TopMargin()
            {
                Width = "72",
                Type = _word_.TableWidthUnitValues.Dxa
            });
            tableCellMarginDefault.AppendChild(new _word_.TableCellLeftMargin()
            {
                Width = 115,
                Type = _word_.TableWidthValues.Dxa
            });
            tableCellMarginDefault.AppendChild(new _word_.BottomMargin()
            {
                Width = "72",
                Type = _word_.TableWidthUnitValues.Dxa
            });
            tableCellMarginDefault.AppendChild(new _word_.TableCellRightMargin()
            {
                Width = 115,
                Type = _word_.TableWidthValues.Dxa
            });

            tableProperties.AppendChild(new _word_.TableLook()
            {
                Val = "04A0",
                FirstRow = true,
                LastRow = false,
                FirstColumn = true,
                LastColumn = false,
                NoHorizontalBand = false,
                NoVerticalBand = true
            });

            _word_.TableGrid tableGrid = table.AppendChild(new _word_.TableGrid());
            tableGrid.AppendChild(new _word_.GridColumn() { Width = "11057" });

            _word_.TableRow tableRow = table.AppendChild(new _word_.TableRow());

            _word_.TableRowProperties tableRowProperties = tableRow.AppendChild(new _word_.TableRowProperties());
            tableRowProperties.AppendChild(new _word_.TableRowHeight() { Val = 64 });

            _word_.TableCell tableCell = tableRow.AppendChild(new _word_.TableCell());

            _word_.TableCellProperties tableCellProperties = tableCell.AppendChild(new _word_.TableCellProperties());
            tableCellProperties.AppendChild(new _word_.TableCellWidth()
            {
                Width = "11058",
                Type = _word_.TableWidthUnitValues.Dxa
            });

            _word_.Paragraph paragraph = tableCell.AppendChild(new _word_.Paragraph());

            _word_.ParagraphProperties paragraphProperties = paragraph.AppendChild(new _word_.ParagraphProperties());
            paragraphProperties.AppendChild(new _word_.SpacingBetweenLines()
            {
                After = "0",
                Line = "0",
                LineRule = _word_.LineSpacingRuleValues.AtLeast
            });
            paragraphProperties.AppendChild(new _word_.Justification() { Val = justificationValue });

            _word_.Run run = paragraph.AppendChild(new _word_.Run());

            _word_.RunProperties runProperties = run.AppendChild(new _word_.RunProperties());
            runProperties.AppendChild(new _word_.RunFonts() { Ascii = "Liberation Sans" });
            runProperties.AppendChild(new _word_.Bold());
            runProperties.AppendChild(new _word_.FontSize() { Val = "20" });

            if (hasPageNum)
            {
                run.AppendChild(new _word_.PageNumber());
                run.AppendChild(new _word_.TabChar());
            }

            else if (hasSectionNum)
            {
                run.AppendChild(new _word_.Text() { Text = Convert.ToString(sectionNum) });
                run.AppendChild(new _word_.TabChar());
            }

            if (text.Length > 0)
                run.AppendChild(new _word_.Text()
                {
                    Text = text,
                    Space = _ooxml_.SpaceProcessingModeValues.Preserve
                });
        }

        public int AddTable(int indexSection, List<TableRow> rows, uint borderTopSize = 4, uint borderLeftSize = 4, uint borderBottomSize = 4, uint borderRightSize = 4, uint borderInsideVerticalSize = 4, uint borderInsideHorizontalSize = 4)
        {
            _word_.Table table = new _word_.Table();

            _word_.TableProperties tableProperties = table.AppendChild(new _word_.TableProperties());
            tableProperties.AppendChild(new _word_.TableStyle() { Val = "TableGrid" });
            tableProperties.AppendChild(new _word_.TableWidth()
            {
                Type = _word_.TableWidthUnitValues.Pct,
                Width = "100%"
            });

            _word_.TableBorders tableBorders = tableProperties.AppendChild(new _word_.TableBorders());
            tableBorders.AppendChild(new _word_.TopBorder()
            {
                Val = _word_.BorderValues.Single,
                Size = borderTopSize,
                Space = 0,
                Color = "auto"
            });
            tableBorders.AppendChild(new _word_.LeftBorder()
            {
                Val = _word_.BorderValues.Single,
                Size = borderLeftSize,
                Space = 0,
                Color = "auto"
            });
            tableBorders.AppendChild(new _word_.BottomBorder()
            {
                Val = _word_.BorderValues.Single,
                Size = borderBottomSize,
                Space = 0,
                Color = "auto"
            });
            tableBorders.AppendChild(new _word_.RightBorder()
            {
                Val = _word_.BorderValues.Single,
                Size = borderRightSize,
                Space = 0,
                Color = "auto"
            });
            tableBorders.AppendChild(new _word_.InsideHorizontalBorder()
            {
                Val = _word_.BorderValues.Single,
                Size = borderInsideHorizontalSize,
                Space = 0,
                Color = "auto"
            });
            tableBorders.AppendChild(new _word_.InsideVerticalBorder()
            {
                Val = _word_.BorderValues.Single,
                Size = borderInsideVerticalSize,
                Space = 0,
                Color = "auto"
            });

            //tableProperties.AppendChild(new _word_.TableLayout() {  Type = _word_.TableLayoutValues.Autofit });

            _word_.TableCellMarginDefault tableCellMarginDefault = tableProperties.AppendChild(new _word_.TableCellMarginDefault());
            tableCellMarginDefault.AppendChild(new _word_.TableCellLeftMargin()
            {
                Type = _word_.TableWidthValues.Dxa,
                Width = 110
            });
            tableCellMarginDefault.AppendChild(new _word_.TableCellRightMargin()
            {
                Type = _word_.TableWidthValues.Dxa,
                Width = 110
            });

            tableProperties.AppendChild(new _word_.TableLook()
            {
                Val = "04a0",
                FirstRow = true,
                LastRow = false,
                FirstColumn = true,
                LastColumn = false,
                NoHorizontalBand = false,
                NoVerticalBand = true
            });

            table.AppendChild(new _word_.TableGrid());

            foreach (TableRow row in rows)
            {
                float cellWidth = (1.0f / row.Cells.Count) * 100;
                string cellWidthString = cellWidth + "%";
                cellWidthString = cellWidthString.Replace(',', '.');

                _word_.TableRow tableRow = table.AppendChild(new _word_.TableRow());

                foreach (TableCell cell in row.Cells)
                {
                    _word_.TableCell tableCell = tableRow.AppendChild(new _word_.TableCell());

                    _word_.TableCellProperties tableCellProperties = tableCell.AppendChild(new _word_.TableCellProperties());
                    tableCellProperties.AppendChild(new _word_.TableCellWidth()
                    {
                        Type = _word_.TableWidthUnitValues.Pct,
                        Width = cellWidthString
                    });
                    tableCellProperties.AppendChild(new _word_.GridSpan() { Val = cell.GridSpan });

                    _word_.TableCellBorders tableCellBorders = tableCellProperties.AppendChild(new _word_.TableCellBorders());
                    if (cell.BorderTopSize > 0)
                        tableCellBorders.AppendChild(new _word_.TopBorder()
                        {
                            Val = _word_.BorderValues.Single,
                            Size = cell.BorderTopSize,
                            Space = 0,
                            Color = "auto"
                        });
                    if (cell.BorderLeftSize > 0)
                        tableCellBorders.AppendChild(new _word_.LeftBorder()
                        {
                            Val = _word_.BorderValues.Single,
                            Size = cell.BorderLeftSize,
                            Space = 0,
                            Color = "auto"
                        });
                    if (cell.BorderBottomSize > 0)
                        tableCellBorders.AppendChild(new _word_.BottomBorder()
                        {
                            Val = _word_.BorderValues.Single,
                            Size = cell.BorderBottomSize,
                            Space = 0,
                            Color = "auto"
                        });
                    if (cell.BorderRightSize > 0)
                        tableCellBorders.AppendChild(new _word_.RightBorder()
                        {
                            Val = _word_.BorderValues.Single,
                            Size = cell.BorderRightSize,
                            Space = 0,
                            Color = "auto"
                        });

                    _word_.Paragraph paragraph = tableCell.AppendChild(new _word_.Paragraph());

                    _word_.ParagraphProperties paragraphProperties = paragraph.AppendChild(new _word_.ParagraphProperties());
                    paragraphProperties.AppendChild(new _word_.SpacingBetweenLines()
                    {
                        After = "0",
                        Line = "0",
                        LineRule = _word_.LineSpacingRuleValues.AtLeast
                    });
                    if (cell.IsCentre)
                        paragraphProperties.AppendChild(new _word_.Justification() { Val = _word_.JustificationValues.Center });

                    _word_.Run run = paragraph.AppendChild(new _word_.Run());

                    _word_.RunProperties runProperties = run.AppendChild(new _word_.RunProperties());
                    if (cell.IsBold)
                        runProperties.AppendChild(new _word_.Bold());

                    run.AppendChild(new _word_.Text
                    {
                        Text = cell.Content,
                        Space = _ooxml_.SpaceProcessingModeValues.Preserve
                    });
                }
            }

            // Add element to section.

            _sections[indexSection].Elements.Add(table);

            return _sections[indexSection].ElementsIndexCurrent++;
        }

        public int AddBarChart(int indexSection, string chartName, List<string> categories, List<Series> series, bool showChartName, bool showValue, ushort gapWidth, sbyte columnOverlap, string categoryAxisName = "", string valueAxisName = "")
        {
            int countCategories = categories.Count;
            int countSeries = series.Count;

            uint idDocElement = _idDocElementNext++;

            string idRelationship = "rId" + _idRelationshipNext++;

            // Add the Paragraph.

            _word_.Paragraph paragraph = new _word_.Paragraph();
            paragraph.AppendChild(new _word_.ParagraphProperties());

            // Add the Run (chart inlined).

            _word_.Run run = paragraph.AppendChild(new _word_.Run());

            _word_.Drawing drawing = run.AppendChild(new _word_.Drawing());

            _drawing_word_.Inline inline = drawing.AppendChild(new _drawing_word_.Inline()
            {
                DistanceFromBottom = 0,
                DistanceFromLeft = 0,
                DistanceFromRight = 0,
                DistanceFromTop = 0
            });
            inline.AppendChild(new _drawing_word_.Extent()
            {
                Cx = 5486400,
                Cy = 3200400
            });
            inline.AppendChild(new _drawing_word_.DocProperties()
            {
                Id = idDocElement,
                Name = chartName,
                Title = chartName
            });

            _drawing_.Graphic graphic = inline.AppendChild(new _drawing_.Graphic());

            _drawing_.GraphicData graphicData = graphic.AppendChild(new _drawing_.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" });
            graphicData.AppendChild(new _drawing_charts_.ChartReference() { Id = idRelationship });

            // Add the ChartPart.

            _package_.ChartPart chartPart = _mainDocumentPart.AddNewPart<_package_.ChartPart>(idRelationship);
            chartPart.ChartSpace = new _drawing_charts_.ChartSpace();
            chartPart.ChartSpace.AppendChild(new _drawing_charts_.Date1904() { Val = false });
            chartPart.ChartSpace.AppendChild(new _drawing_charts_.EditingLanguage() { Val = "en-ZA" });
            chartPart.ChartSpace.AppendChild(new _drawing_charts_.RoundedCorners() { Val = false });

            _ooxml_.AlternateContent alternateContent = chartPart.ChartSpace.AppendChild(new _ooxml_.AlternateContent());

            _ooxml_.AlternateContentChoice alternateContentChoice = alternateContent.AppendChild(new _ooxml_.AlternateContentChoice() { Requires = "c14" });
            alternateContentChoice.AddNamespaceDeclaration("c14", "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
            alternateContentChoice.AppendChild(new _drawing_charts_2010_.Style() { Val = 102 });

            _ooxml_.AlternateContentFallback alternateContentFallback = alternateContent.AppendChild(new _ooxml_.AlternateContentFallback());
            alternateContentFallback.AppendChild(new _drawing_charts_.Style() { Val = 2 });

            // Add the Chart to the ChartSpace.

            _drawing_charts_.Chart chart = chartPart.ChartSpace.AppendChild(new _drawing_charts_.Chart());

            if (showChartName)
            {
                // Add the Title to the Chart.

                _drawing_charts_.Title title = chart.AppendChild(new _drawing_charts_.Title());

                _drawing_charts_.ChartText chartText = title.AppendChild(new _drawing_charts_.ChartText());

                _drawing_charts_.RichText richText = chartText.AppendChild(new _drawing_charts_.RichText());
                richText.AppendChild(new _drawing_.BodyProperties());
                _drawing_.Paragraph paragraphTitle = richText.AppendChild(new _drawing_.Paragraph());
                _drawing_.Run runTitle = paragraphTitle.AppendChild(new _drawing_.Run());
                runTitle.AppendChild(new _drawing_.RunProperties() { Bold = true });
                runTitle.AppendChild(new _drawing_.Text() { Text = chartName });

                title.AppendChild(new _drawing_charts_.Overlay() { Val = false });

                chart.AppendChild(new _drawing_charts_.AutoTitleDeleted() { Val = false });
            }

            // Add the PlotArea to the Chart.

            _drawing_charts_.PlotArea plotArea = chart.AppendChild(new _drawing_charts_.PlotArea());

            // Add the BarChart to the PlotArea.

            _drawing_charts_.BarChart barChart = plotArea.AppendChild(new _drawing_charts_.BarChart());
            barChart.AppendChild(new _drawing_charts_.BarDirection() { Val = _drawing_charts_.BarDirectionValues.Column });
            barChart.AppendChild(new _drawing_charts_.BarGrouping { Val = _drawing_charts_.BarGroupingValues.Clustered });
            barChart.AppendChild(new _drawing_charts_.VaryColors() { Val = false });

            for (int i = 0; i < countSeries; ++i)
            {
                Series seriesCurrent = series[i];

                // Series name.

                _drawing_charts_.BarChartSeries barChartSeries = barChart.AppendChild(new _drawing_charts_.BarChartSeries());
                barChartSeries.AppendChild(new _drawing_charts_.Index() { Val = Convert.ToUInt32(i) });
                barChartSeries.AppendChild(new _drawing_charts_.Order() { Val = Convert.ToUInt32(i) });

                _drawing_charts_.SeriesText seriesText = barChartSeries.AppendChild(new _drawing_charts_.SeriesText());
                seriesText.AppendChild(new _drawing_charts_.NumericValue(seriesCurrent.Name));

                barChartSeries.AppendChild(new _drawing_charts_.InvertIfNegative() { Val = false });

                // Trendline (MovingAverage).

                if (seriesCurrent.TrendlinePeriod >= 2)
                {
                    _drawing_charts_.Trendline trendline = barChartSeries.AppendChild(new _drawing_charts_.Trendline());
                    trendline.AppendChild(new _drawing_charts_.TrendlineType() { Val = _drawing_charts_.TrendlineValues.MovingAverage });
                    trendline.AppendChild(new _drawing_charts_.Period() { Val = seriesCurrent.TrendlinePeriod });
                }

                // CategoryAxis names.

                _drawing_charts_.CategoryAxisData categoryAxisData = barChartSeries.AppendChild(new _drawing_charts_.CategoryAxisData());

                _drawing_charts_.StringLiteral stringLiteral = categoryAxisData.AppendChild(new _drawing_charts_.StringLiteral());
                stringLiteral.AppendChild(new _drawing_charts_.PointCount() { Val = Convert.ToUInt32(countCategories) });

                for (int j = 0; j < countCategories; ++j)
                {
                    _drawing_charts_.StringPoint stringPoint = stringLiteral.AppendChild(new _drawing_charts_.StringPoint() { Index = Convert.ToUInt32(j) });
                    stringPoint.AppendChild(new _drawing_charts_.NumericValue(categories[j]));
                }

                // ValuesAxis data.

                _drawing_charts_.Values values = barChartSeries.AppendChild(new _drawing_charts_.Values());

                _drawing_charts_.NumberLiteral numberLiteral = values.AppendChild(new _drawing_charts_.NumberLiteral());
                numberLiteral.AppendChild(new _drawing_charts_.PointCount() { Val = Convert.ToUInt32(countCategories) });

                for (int j = 0; j < seriesCurrent.Values.Count; ++j)
                {
                    _drawing_charts_.NumericPoint numericPoint = numberLiteral.AppendChild(new _drawing_charts_.NumericPoint() { Index = Convert.ToUInt32(j) });
                    numericPoint.AppendChild(new _drawing_charts_.NumericValue(Convert.ToString(seriesCurrent.Values[j])));
                }
            }

            // DataLabels.

            _drawing_charts_.DataLabels dataLabels = barChart.AppendChild(new _drawing_charts_.DataLabels());
            dataLabels.AppendChild(new _drawing_charts_.ShowLegendKey() { Val = false });
            dataLabels.AppendChild(new _drawing_charts_.ShowValue() { Val = showValue });
            dataLabels.AppendChild(new _drawing_charts_.ShowCategoryName() { Val = false });
            dataLabels.AppendChild(new _drawing_charts_.ShowSeriesName() { Val = false });
            dataLabels.AppendChild(new _drawing_charts_.ShowPercent() { Val = false });
            dataLabels.AppendChild(new _drawing_charts_.ShowBubbleSize() { Val = false });

            // Other BarChart attributes.

            barChart.AppendChild(new _drawing_charts_.GapWidth() { Val = gapWidth });
            barChart.AppendChild(new _drawing_charts_.Overlap() { Val = columnOverlap });
            barChart.AppendChild(new _drawing_charts_.AxisId() { Val = 0 });
            barChart.AppendChild(new _drawing_charts_.AxisId() { Val = 1 });

            // CategoryAxis.

            _drawing_charts_.CategoryAxis categoryAxis = plotArea.AppendChild(new _drawing_charts_.CategoryAxis());
            categoryAxis.AppendChild(new _drawing_charts_.AxisId() { Val = 0 });

            _drawing_charts_.Scaling scaling = categoryAxis.AppendChild(new _drawing_charts_.Scaling());
            scaling.AppendChild(new _drawing_charts_.Orientation() { Val = _drawing_charts_.OrientationValues.MinMax });

            categoryAxis.AppendChild(new _drawing_charts_.Delete() { Val = false });
            categoryAxis.AppendChild(new _drawing_charts_.AxisPosition() { Val = _drawing_charts_.AxisPositionValues.Bottom });

            if (categoryAxisName.Length > 0)
            {
                // Add the Title to the CategoryAxis.

                _drawing_charts_.Title title = categoryAxis.AppendChild(new _drawing_charts_.Title());

                _drawing_charts_.ChartText chartText = title.AppendChild(new _drawing_charts_.ChartText());

                _drawing_charts_.RichText richText = chartText.AppendChild(new _drawing_charts_.RichText());
                _drawing_.BodyProperties bodyProperties = richText.AppendChild(new _drawing_.BodyProperties());
                _drawing_.Paragraph paragraphTitle = richText.AppendChild(new _drawing_.Paragraph());
                _drawing_.Run runTitle = paragraphTitle.AppendChild(new _drawing_.Run());
                runTitle.AppendChild(new _drawing_.RunProperties()
                {
                    Bold = true,
                    FontSize = 1000
                });
                runTitle.AppendChild(new _drawing_.Text() { Text = categoryAxisName });

                title.AppendChild(new _drawing_charts_.Overlay() { Val = false });
            }

            categoryAxis.AppendChild(new _drawing_charts_.NumberingFormat()
            {
                FormatCode = "General",
                SourceLinked = true
            });
            categoryAxis.AppendChild(new _drawing_charts_.MajorTickMark() { Val = _drawing_charts_.TickMarkValues.None });
            categoryAxis.AppendChild(new _drawing_charts_.MinorTickMark() { Val = _drawing_charts_.TickMarkValues.None });
            categoryAxis.AppendChild(new _drawing_charts_.TickLabelPosition() { Val = _drawing_charts_.TickLabelPositionValues.NextTo });
            categoryAxis.AppendChild(new _drawing_charts_.CrossingAxis() { Val = 1 });
            categoryAxis.AppendChild(new _drawing_charts_.Crosses() { Val = _drawing_charts_.CrossesValues.AutoZero });
            categoryAxis.AppendChild(new _drawing_charts_.AutoLabeled() { Val = true });
            categoryAxis.AppendChild(new _drawing_charts_.LabelAlignment() { Val = _drawing_charts_.LabelAlignmentValues.Center });
            categoryAxis.AppendChild(new _drawing_charts_.LabelOffset() { Val = 100 });
            categoryAxis.AppendChild(new _drawing_charts_.NoMultiLevelLabels() { Val = false });

            // ValueAxis.

            _drawing_charts_.ValueAxis valueAxis = plotArea.AppendChild(new _drawing_charts_.ValueAxis());
            valueAxis.AppendChild(new _drawing_charts_.AxisId() { Val = 1 });

            _drawing_charts_.Scaling scalingValue = valueAxis.AppendChild(new _drawing_charts_.Scaling());
            scalingValue.AppendChild(new _drawing_charts_.Orientation() { Val = _drawing_charts_.OrientationValues.MinMax });

            valueAxis.AppendChild(new _drawing_charts_.Delete() { Val = false });
            valueAxis.AppendChild(new _drawing_charts_.AxisPosition() { Val = _drawing_charts_.AxisPositionValues.Left });
            valueAxis.AppendChild(new _drawing_charts_.MajorGridlines());

            if (valueAxisName.Length > 0)
            {
                // Add the Title to the ValueAxis.

                _drawing_charts_.Title title = valueAxis.AppendChild(new _drawing_charts_.Title());

                _drawing_charts_.ChartText chartText = title.AppendChild(new _drawing_charts_.ChartText());

                _drawing_charts_.RichText richText = chartText.AppendChild(new _drawing_charts_.RichText());
                _drawing_.BodyProperties bodyProperties = richText.AppendChild(new _drawing_.BodyProperties());
                _drawing_.Paragraph paragraphTitle = richText.AppendChild(new _drawing_.Paragraph());
                _drawing_.Run runTitle = paragraphTitle.AppendChild(new _drawing_.Run());
                runTitle.AppendChild(new _drawing_.RunProperties()
                {
                    Bold = true,
                    FontSize = 1000
                });
                runTitle.AppendChild(new _drawing_.Text() { Text = valueAxisName });

                title.AppendChild(new _drawing_charts_.Overlay() { Val = false });
            }

            valueAxis.AppendChild(new _drawing_charts_.NumberingFormat()
            {
                FormatCode = "General",
                SourceLinked = true
            });
            valueAxis.AppendChild(new _drawing_charts_.MajorTickMark() { Val = _drawing_charts_.TickMarkValues.Outside });
            valueAxis.AppendChild(new _drawing_charts_.MinorTickMark() { Val = _drawing_charts_.TickMarkValues.None });
            valueAxis.AppendChild(new _drawing_charts_.TickLabelPosition() { Val = _drawing_charts_.TickLabelPositionValues.NextTo });
            valueAxis.AppendChild(new _drawing_charts_.CrossingAxis() { Val = 0 });
            valueAxis.AppendChild(new _drawing_charts_.Crosses() { Val = _drawing_charts_.CrossesValues.AutoZero });
            valueAxis.AppendChild(new _drawing_charts_.CrossBetween() { Val = _drawing_charts_.CrossBetweenValues.Between });

            // Legend.

            _drawing_charts_.Legend legend = chart.AppendChild(new _drawing_charts_.Legend());
            legend.AppendChild(new _drawing_charts_.LegendPosition() { Val = _drawing_charts_.LegendPositionValues.Bottom });
            legend.AppendChild(new _drawing_charts_.Overlay() { Val = false });

            // Other Chart attributes.

            chart.AppendChild(new _drawing_charts_.PlotVisibleOnly() { Val = true });
            chart.AppendChild(new _drawing_charts_.DisplayBlanksAs() { Val = _drawing_charts_.DisplayBlanksAsValues.Gap });
            chart.AppendChild(new _drawing_charts_.ShowDataLabelsOverMaximum() { Val = false });

            // Add element to section.

            _sections[indexSection].Elements.Add(paragraph);

            return _sections[indexSection].ElementsIndexCurrent++;
        }

        public int AddLineChart(int indexSection, string chartName, List<string> categories, List<Series> series, bool showChartName, bool showValue, bool showLegend, string categoryAxisName = "", string valueAxisName = "")
        {
            int countCategories = categories.Count;
            int countSeries = series.Count;

            uint idDocElement = _idDocElementNext++;

            string idRelationship = "rId" + _idRelationshipNext++;

            // Add the Paragraph.

            _word_.Paragraph paragraph = new _word_.Paragraph();
            paragraph.AppendChild(new _word_.ParagraphProperties());

            // Add the Run (chart inlined).

            _word_.Run run = paragraph.AppendChild(new _word_.Run());

            _word_.Drawing drawing = run.AppendChild(new _word_.Drawing());

            _drawing_word_.Inline inline = drawing.AppendChild(new _drawing_word_.Inline()
            {
                DistanceFromBottom = 0,
                DistanceFromLeft = 0,
                DistanceFromRight = 0,
                DistanceFromTop = 0
            });
            inline.AppendChild(new _drawing_word_.Extent()
            {
                Cx = 5486400,
                Cy = 3200400
            });
            inline.AppendChild(new _drawing_word_.DocProperties()
            {
                Id = idDocElement,
                Name = chartName,
                Title = chartName
            });

            _drawing_.Graphic graphic = inline.AppendChild(new _drawing_.Graphic());

            _drawing_.GraphicData graphicData = graphic.AppendChild(new _drawing_.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" });
            graphicData.AppendChild(new _drawing_charts_.ChartReference() { Id = idRelationship });

            // Add the ChartPart.

            _package_.ChartPart chartPart = _mainDocumentPart.AddNewPart<_package_.ChartPart>(idRelationship);
            chartPart.ChartSpace = new _drawing_charts_.ChartSpace();
            chartPart.ChartSpace.AppendChild(new _drawing_charts_.Date1904() { Val = false });
            chartPart.ChartSpace.AppendChild(new _drawing_charts_.EditingLanguage() { Val = "en-ZA" });
            chartPart.ChartSpace.AppendChild(new _drawing_charts_.RoundedCorners() { Val = false });

            _ooxml_.AlternateContent alternateContent = chartPart.ChartSpace.AppendChild(new _ooxml_.AlternateContent());

            _ooxml_.AlternateContentChoice alternateContentChoice = alternateContent.AppendChild(new _ooxml_.AlternateContentChoice() { Requires = "c14" });
            alternateContentChoice.AddNamespaceDeclaration("c14", "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
            alternateContentChoice.AppendChild(new _drawing_charts_2010_.Style() { Val = 102 });

            _ooxml_.AlternateContentFallback alternateContentFallback = alternateContent.AppendChild(new _ooxml_.AlternateContentFallback());
            alternateContentFallback.AppendChild(new _drawing_charts_.Style() { Val = 2 });

            // Add the Chart to the ChartSpace.

            _drawing_charts_.Chart chart = chartPart.ChartSpace.AppendChild(new _drawing_charts_.Chart());

            if (showChartName)
            {
                // Add the Title to the Chart.

                _drawing_charts_.Title title = chart.AppendChild(new _drawing_charts_.Title());

                _drawing_charts_.ChartText chartText = title.AppendChild(new _drawing_charts_.ChartText());

                _drawing_charts_.RichText richText = chartText.AppendChild(new _drawing_charts_.RichText());
                richText.AppendChild(new _drawing_.BodyProperties());
                _drawing_.Paragraph paragraphTitle = richText.AppendChild(new _drawing_.Paragraph());
                _drawing_.Run runTitle = paragraphTitle.AppendChild(new _drawing_.Run());
                runTitle.AppendChild(new _drawing_.RunProperties() { Bold = true });
                runTitle.AppendChild(new _drawing_.Text() { Text = chartName });

                title.AppendChild(new _drawing_charts_.Overlay() { Val = false });

                chart.AppendChild(new _drawing_charts_.AutoTitleDeleted() { Val = false });
            }

            // Add the PlotArea to the Chart.

            _drawing_charts_.PlotArea plotArea = chart.AppendChild(new _drawing_charts_.PlotArea());

            // Add the LineChart to the PlotArea.

            _drawing_charts_.LineChart lineChart = plotArea.AppendChild(new _drawing_charts_.LineChart());
            lineChart.AppendChild(new _drawing_charts_.Grouping() { Val = _drawing_charts_.GroupingValues.Standard });
            lineChart.AppendChild(new _drawing_charts_.VaryColors() { Val = false });


            for (int i = 0; i < countSeries; ++i)
            {
                Series seriesCurrent = series[i];

                // Series name.

                _drawing_charts_.LineChartSeries lineChartSeries = lineChart.AppendChild(new _drawing_charts_.LineChartSeries());
                lineChartSeries.AppendChild(new _drawing_charts_.Index() { Val = Convert.ToUInt32(i) });
                lineChartSeries.AppendChild(new _drawing_charts_.Order() { Val = Convert.ToUInt32(i) });

                _drawing_charts_.SeriesText seriesText = lineChartSeries.AppendChild(new _drawing_charts_.SeriesText());
                seriesText.AppendChild(new _drawing_charts_.NumericValue(seriesCurrent.Name));

                // Marker symbol.

                _drawing_charts_.Marker marker = lineChartSeries.AppendChild(new _drawing_charts_.Marker());
                marker.AppendChild(new _drawing_charts_.Symbol() { Val = _drawing_charts_.MarkerStyleValues.Circle });

                // CategoryAxis names.

                _drawing_charts_.CategoryAxisData categoryAxisData = lineChartSeries.AppendChild(new _drawing_charts_.CategoryAxisData());

                _drawing_charts_.StringLiteral stringLiteral = categoryAxisData.AppendChild(new _drawing_charts_.StringLiteral());
                stringLiteral.AppendChild(new _drawing_charts_.PointCount() { Val = Convert.ToUInt32(countCategories) });

                for (int j = 0; j < countCategories; ++j)
                {
                    _drawing_charts_.StringPoint stringPoint = stringLiteral.AppendChild(new _drawing_charts_.StringPoint() { Index = Convert.ToUInt32(j) });
                    stringPoint.AppendChild(new _drawing_charts_.NumericValue(categories[j]));
                }

                // ValuesAxis data.

                _drawing_charts_.Values values = lineChartSeries.AppendChild(new _drawing_charts_.Values());

                _drawing_charts_.NumberLiteral numberLiteral = values.AppendChild(new _drawing_charts_.NumberLiteral());
                numberLiteral.AppendChild(new _drawing_charts_.PointCount() { Val = Convert.ToUInt32(countCategories) });

                for (int j = 0; j < seriesCurrent.Values.Count; ++j)
                {
                    _drawing_charts_.NumericPoint numericPoint = numberLiteral.AppendChild(new _drawing_charts_.NumericPoint() { Index = Convert.ToUInt32(j) });
                    numericPoint.AppendChild(new _drawing_charts_.NumericValue(Convert.ToString(seriesCurrent.Values[j])));
                }

                lineChartSeries.AppendChild(new _drawing_charts_.Smooth() { Val = false });
            }

            // DataLabels.

            _drawing_charts_.DataLabels dataLabels = lineChart.AppendChild(new _drawing_charts_.DataLabels());
            dataLabels.AppendChild(new _drawing_charts_.ShowLegendKey() { Val = false });
            dataLabels.AppendChild(new _drawing_charts_.ShowValue() { Val = showValue });
            dataLabels.AppendChild(new _drawing_charts_.ShowCategoryName() { Val = false });
            dataLabels.AppendChild(new _drawing_charts_.ShowSeriesName() { Val = false });
            dataLabels.AppendChild(new _drawing_charts_.ShowPercent() { Val = false });
            dataLabels.AppendChild(new _drawing_charts_.ShowBubbleSize() { Val = false });

            // Other LineChart attributes.

            lineChart.AppendChild(new _drawing_charts_.ShowMarker() { Val = true });
            lineChart.AppendChild(new _drawing_charts_.Smooth() { Val = false });
            lineChart.AppendChild(new _drawing_charts_.AxisId() { Val = 0 });
            lineChart.AppendChild(new _drawing_charts_.AxisId() { Val = 1 });

            // CategoryAxis.

            _drawing_charts_.CategoryAxis categoryAxis = plotArea.AppendChild(new _drawing_charts_.CategoryAxis());
            categoryAxis.AppendChild(new _drawing_charts_.AxisId() { Val = 0 });

            _drawing_charts_.Scaling scaling = categoryAxis.AppendChild(new _drawing_charts_.Scaling());
            scaling.AppendChild(new _drawing_charts_.Orientation() { Val = _drawing_charts_.OrientationValues.MinMax });

            categoryAxis.AppendChild(new _drawing_charts_.Delete() { Val = false });
            categoryAxis.AppendChild(new _drawing_charts_.AxisPosition() { Val = _drawing_charts_.AxisPositionValues.Bottom });

            if (categoryAxisName.Length > 0)
            {
                // Add the Title to the CategoryAxis.

                _drawing_charts_.Title title = categoryAxis.AppendChild(new _drawing_charts_.Title());

                _drawing_charts_.ChartText chartText = title.AppendChild(new _drawing_charts_.ChartText());

                _drawing_charts_.RichText richText = chartText.AppendChild(new _drawing_charts_.RichText());
                _drawing_.BodyProperties bodyProperties = richText.AppendChild(new _drawing_.BodyProperties());
                _drawing_.Paragraph paragraphTitle = richText.AppendChild(new _drawing_.Paragraph());
                _drawing_.Run runTitle = paragraphTitle.AppendChild(new _drawing_.Run());
                runTitle.AppendChild(new _drawing_.RunProperties()
                {
                    Bold = true,
                    FontSize = 1000
                });
                runTitle.AppendChild(new _drawing_.Text() { Text = categoryAxisName });

                title.AppendChild(new _drawing_charts_.Overlay() { Val = false });
            }

            categoryAxis.AppendChild(new _drawing_charts_.NumberingFormat()
            {
                FormatCode = "General",
                SourceLinked = true
            });
            categoryAxis.AppendChild(new _drawing_charts_.MajorTickMark() { Val = _drawing_charts_.TickMarkValues.None });
            categoryAxis.AppendChild(new _drawing_charts_.MinorTickMark() { Val = _drawing_charts_.TickMarkValues.None });
            categoryAxis.AppendChild(new _drawing_charts_.TickLabelPosition() { Val = _drawing_charts_.TickLabelPositionValues.NextTo });
            categoryAxis.AppendChild(new _drawing_charts_.CrossingAxis() { Val = 1 });
            categoryAxis.AppendChild(new _drawing_charts_.Crosses() { Val = _drawing_charts_.CrossesValues.AutoZero });
            categoryAxis.AppendChild(new _drawing_charts_.AutoLabeled() { Val = true });
            categoryAxis.AppendChild(new _drawing_charts_.LabelAlignment() { Val = _drawing_charts_.LabelAlignmentValues.Center });
            categoryAxis.AppendChild(new _drawing_charts_.LabelOffset() { Val = 100 });
            categoryAxis.AppendChild(new _drawing_charts_.NoMultiLevelLabels() { Val = false });

            // ValueAxis.

            _drawing_charts_.ValueAxis valueAxis = plotArea.AppendChild(new _drawing_charts_.ValueAxis());
            valueAxis.AppendChild(new _drawing_charts_.AxisId() { Val = 1 });

            _drawing_charts_.Scaling scalingValue = valueAxis.AppendChild(new _drawing_charts_.Scaling());
            scalingValue.AppendChild(new _drawing_charts_.Orientation() { Val = _drawing_charts_.OrientationValues.MinMax });

            valueAxis.AppendChild(new _drawing_charts_.Delete() { Val = false });
            valueAxis.AppendChild(new _drawing_charts_.AxisPosition() { Val = _drawing_charts_.AxisPositionValues.Left });
            valueAxis.AppendChild(new _drawing_charts_.MajorGridlines());

            if (valueAxisName.Length > 0)
            {
                // Add the Title to the ValueAxis.

                _drawing_charts_.Title title = valueAxis.AppendChild(new _drawing_charts_.Title());

                _drawing_charts_.ChartText chartText = title.AppendChild(new _drawing_charts_.ChartText());

                _drawing_charts_.RichText richText = chartText.AppendChild(new _drawing_charts_.RichText());
                _drawing_.BodyProperties bodyProperties = richText.AppendChild(new _drawing_.BodyProperties());
                _drawing_.Paragraph paragraphTitle = richText.AppendChild(new _drawing_.Paragraph());
                _drawing_.Run runTitle = paragraphTitle.AppendChild(new _drawing_.Run());
                runTitle.AppendChild(new _drawing_.RunProperties()
                {
                    Bold = true,
                    FontSize = 1000
                });
                runTitle.AppendChild(new _drawing_.Text() { Text = valueAxisName });

                title.AppendChild(new _drawing_charts_.Overlay() { Val = false });
            }

            valueAxis.AppendChild(new _drawing_charts_.NumberingFormat()
            {
                FormatCode = "General",
                SourceLinked = true
            });
            valueAxis.AppendChild(new _drawing_charts_.MajorTickMark() { Val = _drawing_charts_.TickMarkValues.Outside });
            valueAxis.AppendChild(new _drawing_charts_.MinorTickMark() { Val = _drawing_charts_.TickMarkValues.None });
            valueAxis.AppendChild(new _drawing_charts_.TickLabelPosition() { Val = _drawing_charts_.TickLabelPositionValues.NextTo });
            valueAxis.AppendChild(new _drawing_charts_.CrossingAxis() { Val = 0 });
            valueAxis.AppendChild(new _drawing_charts_.Crosses() { Val = _drawing_charts_.CrossesValues.AutoZero });
            valueAxis.AppendChild(new _drawing_charts_.CrossBetween() { Val = _drawing_charts_.CrossBetweenValues.Between });

            // Legend.

            if (showLegend)
            {
                _drawing_charts_.Legend legend = chart.AppendChild(new _drawing_charts_.Legend());
                legend.AppendChild(new _drawing_charts_.LegendPosition() { Val = _drawing_charts_.LegendPositionValues.Bottom });
                legend.AppendChild(new _drawing_charts_.Overlay() { Val = false });
            }

            // Other Chart attributes.

            chart.AppendChild(new _drawing_charts_.PlotVisibleOnly() { Val = true });
            chart.AppendChild(new _drawing_charts_.DisplayBlanksAs() { Val = _drawing_charts_.DisplayBlanksAsValues.Gap });
            chart.AppendChild(new _drawing_charts_.ShowDataLabelsOverMaximum() { Val = false });

            // Add element to section.

            _sections[indexSection].Elements.Add(paragraph);

            return _sections[indexSection].ElementsIndexCurrent++;
        }

        public int AddScatterChart(int indexSection, string chartName, List<Series_Point> series, bool showChartName, bool showValue, bool showLegend, double xAxisMin, double xAxisMax, UnitType xAxisUnit = UnitType.None, UnitType yAxisUnit = UnitType.Temperature_Celsius, string xAxisName = "", string yAxisName = "")
        {
            int countSeries = series.Count;

            uint idDocElement = _idDocElementNext++;

            string idRelationship = "rId" + _idRelationshipNext++;

            // Add the Paragraph.

            _word_.Paragraph paragraph = new _word_.Paragraph();
            paragraph.AppendChild(new _word_.ParagraphProperties());

            // Add the Run (chart inlined).

            _word_.Run run = paragraph.AppendChild(new _word_.Run());

            _word_.Drawing drawing = run.AppendChild(new _word_.Drawing());

            _drawing_word_.Inline inline = drawing.AppendChild(new _drawing_word_.Inline()
            {
                DistanceFromBottom = 0,
                DistanceFromLeft = 0,
                DistanceFromRight = 0,
                DistanceFromTop = 0
            });
            inline.AppendChild(new _drawing_word_.Extent()
            {
                Cx = 5486400,
                Cy = 3200400
            });
            inline.AppendChild(new _drawing_word_.DocProperties()
            {
                Id = idDocElement,
                Name = chartName,
                Title = chartName
            });

            _drawing_.Graphic graphic = inline.AppendChild(new _drawing_.Graphic());

            _drawing_.GraphicData graphicData = graphic.AppendChild(new _drawing_.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" });
            graphicData.AppendChild(new _drawing_charts_.ChartReference() { Id = idRelationship });

            // Add the ChartPart.

            _package_.ChartPart chartPart = _mainDocumentPart.AddNewPart<_package_.ChartPart>(idRelationship);
            chartPart.ChartSpace = new _drawing_charts_.ChartSpace();
            chartPart.ChartSpace.AppendChild(new _drawing_charts_.Date1904() { Val = false });
            chartPart.ChartSpace.AppendChild(new _drawing_charts_.EditingLanguage() { Val = "en-ZA" });
            chartPart.ChartSpace.AppendChild(new _drawing_charts_.RoundedCorners() { Val = false });

            _ooxml_.AlternateContent alternateContent = chartPart.ChartSpace.AppendChild(new _ooxml_.AlternateContent());

            _ooxml_.AlternateContentChoice alternateContentChoice = alternateContent.AppendChild(new _ooxml_.AlternateContentChoice() { Requires = "c14" });
            alternateContentChoice.AddNamespaceDeclaration("c14", "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
            alternateContentChoice.AppendChild(new _drawing_charts_2010_.Style() { Val = 102 });

            _ooxml_.AlternateContentFallback alternateContentFallback = alternateContent.AppendChild(new _ooxml_.AlternateContentFallback());
            alternateContentFallback.AppendChild(new _drawing_charts_.Style() { Val = 2 });

            // Add the Chart to the ChartSpace.

            _drawing_charts_.Chart chart = chartPart.ChartSpace.AppendChild(new _drawing_charts_.Chart());

            if (showChartName)
            {
                // Add the Title to the Chart.

                _drawing_charts_.Title title = chart.AppendChild(new _drawing_charts_.Title());

                _drawing_charts_.ChartText chartText = title.AppendChild(new _drawing_charts_.ChartText());

                _drawing_charts_.RichText richText = chartText.AppendChild(new _drawing_charts_.RichText());
                richText.AppendChild(new _drawing_.BodyProperties());
                _drawing_.Paragraph paragraphTitle = richText.AppendChild(new _drawing_.Paragraph());
                _drawing_.Run runTitle = paragraphTitle.AppendChild(new _drawing_.Run());
                runTitle.AppendChild(new _drawing_.RunProperties() { Bold = true });
                runTitle.AppendChild(new _drawing_.Text() { Text = chartName });

                title.AppendChild(new _drawing_charts_.Overlay() { Val = false });

                chart.AppendChild(new _drawing_charts_.AutoTitleDeleted() { Val = false });
            }

            // Add the PlotArea to the Chart.

            _drawing_charts_.PlotArea plotArea = chart.AppendChild(new _drawing_charts_.PlotArea());

            // Add the ScatterChart to the PlotArea.

            _drawing_charts_.ScatterChart scatterChart = plotArea.AppendChild(new _drawing_charts_.ScatterChart());

            _drawing_charts_.ScatterStyle scatterstyle = scatterChart.AppendChild(new _drawing_charts_.ScatterStyle() { Val = _drawing_charts_.ScatterStyleValues.Line });

            scatterChart.AppendChild(new _drawing_charts_.VaryColors() { Val = false });

            for (int i = 0; i < countSeries; ++i)
            {
                Series_Point seriesCurrent = series[i];
                int seriesCurrentCount = seriesCurrent.Values.Count;

                // Series name.

                _drawing_charts_.ScatterChartSeries scatterChartSeries = scatterChart.AppendChild(new _drawing_charts_.ScatterChartSeries());
                scatterChartSeries.AppendChild(new _drawing_charts_.Index() { Val = Convert.ToUInt32(i) });
                scatterChartSeries.AppendChild(new _drawing_charts_.Order() { Val = Convert.ToUInt32(i) });

                _drawing_charts_.SeriesText seriesText = scatterChartSeries.AppendChild(new _drawing_charts_.SeriesText());
                seriesText.AppendChild(new _drawing_charts_.NumericValue(seriesCurrent.Name));

                _drawing_charts_.ChartShapeProperties chartShapeProperties = scatterChartSeries.AppendChild(new _drawing_charts_.ChartShapeProperties());

                {
                    _drawing_.SolidFill solidFill = chartShapeProperties.AppendChild(new _drawing_.SolidFill());
                    solidFill.AppendChild(new _drawing_.RgbColorModelHex() { Val = new _ooxml_.HexBinaryValue("004586") });
                }

                _drawing_.Outline outline = chartShapeProperties.AppendChild(new _drawing_.Outline() { Width = 3600 });

                {
                    _drawing_.SolidFill solidFill = outline.AppendChild(new _drawing_.SolidFill());
                    solidFill.AppendChild(new _drawing_.RgbColorModelHex() { Val = new _ooxml_.HexBinaryValue("004586") });

                    outline.AppendChild(new _drawing_.Round());
                }


                // Marker symbol.

                _drawing_charts_.Marker marker = scatterChartSeries.AppendChild(new _drawing_charts_.Marker());
                marker.AppendChild(new _drawing_charts_.Symbol() { Val = _drawing_charts_.MarkerStyleValues.None });

                // X Values.

                {
                    _drawing_charts_.XValues xvalues = scatterChartSeries.AppendChild(new _drawing_charts_.XValues());

                    _drawing_charts_.NumberLiteral numberLiteral = xvalues.AppendChild(new _drawing_charts_.NumberLiteral());
                    numberLiteral.AppendChild(new _drawing_charts_.PointCount() { Val = Convert.ToUInt32(seriesCurrentCount) });

                    for (int j = 0; j < seriesCurrentCount; ++j)
                    {
                        _drawing_charts_.NumericPoint numericPoint = numberLiteral.AppendChild(new _drawing_charts_.NumericPoint() { Index = Convert.ToUInt32(j) });
                        numericPoint.AppendChild(new _drawing_charts_.NumericValue(Convert.ToString(seriesCurrent.Values[j].Key).Replace(',', '.')));
                    }
                }

                // Y values.

                {
                    _drawing_charts_.YValues yvalues = scatterChartSeries.AppendChild(new _drawing_charts_.YValues());

                    _drawing_charts_.NumberLiteral numberLiteral = yvalues.AppendChild(new _drawing_charts_.NumberLiteral());
                    numberLiteral.AppendChild(new _drawing_charts_.PointCount() { Val = Convert.ToUInt32(seriesCurrentCount) });

                    for (int j = 0; j < seriesCurrentCount; ++j)
                    {
                        _drawing_charts_.NumericPoint numericPoint = numberLiteral.AppendChild(new _drawing_charts_.NumericPoint() { Index = Convert.ToUInt32(j) });
                        numericPoint.AppendChild(new _drawing_charts_.NumericValue(Convert.ToString(seriesCurrent.Values[j].Value).Replace(',', '.')));
                    }
                }

                scatterChartSeries.AppendChild(new _drawing_charts_.Smooth() { Val = false });
            }

            // DataLabels.

            {
                _drawing_charts_.DataLabels dataLabels = scatterChart.AppendChild(new _drawing_charts_.DataLabels());
                dataLabels.AppendChild(new _drawing_charts_.ShowLegendKey() { Val = false });
                dataLabels.AppendChild(new _drawing_charts_.ShowValue() { Val = showValue });
                dataLabels.AppendChild(new _drawing_charts_.ShowCategoryName() { Val = false });
                dataLabels.AppendChild(new _drawing_charts_.ShowSeriesName() { Val = false });
                dataLabels.AppendChild(new _drawing_charts_.ShowPercent() { Val = false });
                dataLabels.AppendChild(new _drawing_charts_.ShowLeaderLines() { Val = false });
            }

            // Other LineChart attributes.

            scatterChart.AppendChild(new _drawing_charts_.AxisId() { Val = 0 });
            scatterChart.AppendChild(new _drawing_charts_.AxisId() { Val = 1 });

            // X Axis.

            {
                _drawing_charts_.ValueAxis valueAxis = plotArea.AppendChild(new _drawing_charts_.ValueAxis());
                valueAxis.AppendChild(new _drawing_charts_.AxisId() { Val = 0 });

                _drawing_charts_.Scaling scaling = valueAxis.AppendChild(new _drawing_charts_.Scaling());
                scaling.AppendChild(new _drawing_charts_.Orientation() { Val = _drawing_charts_.OrientationValues.MinMax });
                scaling.AppendChild(new _drawing_charts_.MaxAxisValue() { Val = xAxisMax });
                scaling.AppendChild(new _drawing_charts_.MinAxisValue() { Val = xAxisMin });

                valueAxis.AppendChild(new _drawing_charts_.Delete() { Val = false });
                valueAxis.AppendChild(new _drawing_charts_.AxisPosition() { Val = _drawing_charts_.AxisPositionValues.Bottom });

                if (xAxisName.Length > 0)
                {
                    // Add the Title to the X Axis.

                    _drawing_charts_.Title title = valueAxis.AppendChild(new _drawing_charts_.Title());

                    _drawing_charts_.ChartText chartText = title.AppendChild(new _drawing_charts_.ChartText());

                    _drawing_charts_.RichText richText = chartText.AppendChild(new _drawing_charts_.RichText());
                    _drawing_.BodyProperties bodyProperties = richText.AppendChild(new _drawing_.BodyProperties());
                    _drawing_.Paragraph paragraphTitle = richText.AppendChild(new _drawing_.Paragraph());
                    {
                        _drawing_.Run runTitle = paragraphTitle.AppendChild(new _drawing_.Run());
                        runTitle.AppendChild(new _drawing_.RunProperties()
                        {
                            Bold = true,
                            FontSize = 1000
                        });
                        runTitle.AppendChild(new _drawing_.Text() { Text = xAxisName });
                    }

                    if (xAxisUnit == UnitType.Temperature_Celsius)
                    {
                        _drawing_.Run runTitle = paragraphTitle.AppendChild(new _drawing_.Run());
                        runTitle.AppendChild(new _drawing_.RunProperties()
                        {
                            Bold = true,
                            FontSize = 1000
                        });
                        runTitle.AppendChild(new _drawing_.Text() { Text = " (\u00B0C)" });
                    }

                    title.AppendChild(new _drawing_charts_.Overlay() { Val = false });
                }

                valueAxis.AppendChild(new _drawing_charts_.NumberingFormat()
                {
                    FormatCode = "YY/MM/DD HH:mm",
                    SourceLinked = true
                });
                valueAxis.AppendChild(new _drawing_charts_.MajorTickMark() { Val = _drawing_charts_.TickMarkValues.Outside });
                valueAxis.AppendChild(new _drawing_charts_.MinorTickMark() { Val = _drawing_charts_.TickMarkValues.None });
                valueAxis.AppendChild(new _drawing_charts_.TickLabelPosition() { Val = _drawing_charts_.TickLabelPositionValues.Low });

                _drawing_charts_.ChartShapeProperties chartShapeProperties = valueAxis.AppendChild(new _drawing_charts_.ChartShapeProperties());
                {
                    _drawing_.Outline outline = chartShapeProperties.AppendChild(new _drawing_.Outline() { Width = 18000 });
                    {
                        _drawing_.SolidFill solidfill = outline.AppendChild(new _drawing_.SolidFill());
                        solidfill.AppendChild(new _drawing_.RgbColorModelHex() { Val = new _ooxml_.HexBinaryValue("b3b3b3") });

                        outline.AppendChild(new _drawing_.Round());
                    }
                }

                valueAxis.AppendChild(new _drawing_charts_.CrossingAxis() { Val = 1 });
                valueAxis.AppendChild(new _drawing_charts_.Crosses() { Val = _drawing_charts_.CrossesValues.AutoZero });
                valueAxis.AppendChild(new _drawing_charts_.CrossBetween() { Val = _drawing_charts_.CrossBetweenValues.MidpointCategory });
            }

            // Y Axis.

            {
                _drawing_charts_.ValueAxis valueAxis = plotArea.AppendChild(new _drawing_charts_.ValueAxis());
                valueAxis.AppendChild(new _drawing_charts_.AxisId() { Val = 1 });

                _drawing_charts_.Scaling scalingValue = valueAxis.AppendChild(new _drawing_charts_.Scaling());
                scalingValue.AppendChild(new _drawing_charts_.Orientation() { Val = _drawing_charts_.OrientationValues.MinMax });

                valueAxis.AppendChild(new _drawing_charts_.Delete() { Val = false });
                valueAxis.AppendChild(new _drawing_charts_.AxisPosition() { Val = _drawing_charts_.AxisPositionValues.Left });
                _drawing_charts_.MajorGridlines majorGridines = valueAxis.AppendChild(new _drawing_charts_.MajorGridlines());

                {
                    _drawing_charts_.ChartShapeProperties chartShapeProperties = majorGridines.AppendChild(new _drawing_charts_.ChartShapeProperties());
                    {
                        _drawing_.Outline outline = chartShapeProperties.AppendChild(new _drawing_.Outline());
                        {
                            _drawing_.SolidFill solidfill = outline.AppendChild(new _drawing_.SolidFill());
                            solidfill.AppendChild(new _drawing_.RgbColorModelHex() { Val = new _ooxml_.HexBinaryValue("b3b3b3") });
                        }
                    }
                }

                if (yAxisName.Length > 0)
                {
                    // Add the Title to the ValueAxis.

                    _drawing_charts_.Title title = valueAxis.AppendChild(new _drawing_charts_.Title());

                    _drawing_charts_.ChartText chartText = title.AppendChild(new _drawing_charts_.ChartText());

                    _drawing_charts_.RichText richText = chartText.AppendChild(new _drawing_charts_.RichText());
                    _drawing_.BodyProperties bodyProperties = richText.AppendChild(new _drawing_.BodyProperties());
                    _drawing_.Paragraph paragraphTitle = richText.AppendChild(new _drawing_.Paragraph());
                    {
                        _drawing_.Run runTitle = paragraphTitle.AppendChild(new _drawing_.Run());
                        runTitle.AppendChild(new _drawing_.RunProperties()
                        {
                            Bold = true,
                            FontSize = 1000
                        });
                        runTitle.AppendChild(new _drawing_.Text() { Text = yAxisName });
                    }

                    if (yAxisUnit == UnitType.Temperature_Celsius)
                    {
                        _drawing_.Run runTitle = paragraphTitle.AppendChild(new _drawing_.Run());
                        runTitle.AppendChild(new _drawing_.RunProperties()
                        {
                            Bold = true,
                            FontSize = 1000
                        });
                        runTitle.AppendChild(new _drawing_.Text() { Text = " (\u00B0C)" });
                    }

                    title.AppendChild(new _drawing_charts_.Overlay() { Val = false });
                }

                valueAxis.AppendChild(new _drawing_charts_.NumberingFormat()
                {
                    FormatCode = "General",
                    SourceLinked = true
                });
                valueAxis.AppendChild(new _drawing_charts_.MajorTickMark() { Val = _drawing_charts_.TickMarkValues.Outside });
                valueAxis.AppendChild(new _drawing_charts_.MinorTickMark() { Val = _drawing_charts_.TickMarkValues.None });
                valueAxis.AppendChild(new _drawing_charts_.TickLabelPosition() { Val = _drawing_charts_.TickLabelPositionValues.NextTo });

                {
                    _drawing_charts_.ChartShapeProperties chartShapeProperties = valueAxis.AppendChild(new _drawing_charts_.ChartShapeProperties());
                    {
                        _drawing_.Outline outline = chartShapeProperties.AppendChild(new _drawing_.Outline() { Width = 18000 });
                        {
                            _drawing_.SolidFill solidfill = outline.AppendChild(new _drawing_.SolidFill());
                            solidfill.AppendChild(new _drawing_.RgbColorModelHex() { Val = new _ooxml_.HexBinaryValue("b3b3b3") });

                            outline.AppendChild(new _drawing_.Round());
                        }
                    }
                }

                valueAxis.AppendChild(new _drawing_charts_.CrossingAxis() { Val = 0 });
                valueAxis.AppendChild(new _drawing_charts_.Crosses() { Val = _drawing_charts_.CrossesValues.AutoZero });
                valueAxis.AppendChild(new _drawing_charts_.CrossBetween() { Val = _drawing_charts_.CrossBetweenValues.MidpointCategory });
            }

            {
                _drawing_charts_.ShapeProperties shapeProperties = plotArea.AppendChild(new _drawing_charts_.ShapeProperties());
                {
                    shapeProperties.AppendChild(new _drawing_.NoFill());

                    _drawing_.Outline outline = shapeProperties.AppendChild(new _drawing_.Outline());
                    {
                        _drawing_.SolidFill solidfill = outline.AppendChild(new _drawing_.SolidFill());
                        solidfill.AppendChild(new _drawing_.RgbColorModelHex() { Val = new _ooxml_.HexBinaryValue("b3b3b3") });
                    }
                }
            }

            // Legend.

            if (showLegend)
            {
                _drawing_charts_.Legend legend = chart.AppendChild(new _drawing_charts_.Legend());
                legend.AppendChild(new _drawing_charts_.LegendPosition() { Val = _drawing_charts_.LegendPositionValues.Bottom });
                legend.AppendChild(new _drawing_charts_.Overlay() { Val = false });
            }

            // Other Chart attributes.

            chart.AppendChild(new _drawing_charts_.PlotVisibleOnly() { Val = true });
            chart.AppendChild(new _drawing_charts_.DisplayBlanksAs() { Val = _drawing_charts_.DisplayBlanksAsValues.Span });

            // Add element to section.

            _sections[indexSection].Elements.Add(paragraph);

            return _sections[indexSection].ElementsIndexCurrent++;
        }

        public int AddAnchoredPicture(int indexSection, string filePath, uint imageWidth, uint imageHeight, AlignmentHorizontal alignment = AlignmentHorizontal.Left, uint positionY = 0)
        {
            string alignmentString = "left";

            if (alignment == AlignmentHorizontal.Centre)
                alignmentString = "center";

            else if (alignment == AlignmentHorizontal.Right)
                alignmentString = "right";

            uint idDocElement = _idDocElementNext++;
            uint idPicture = _idPictureNext++;

            string idRelationship = "rId" + _idRelationshipNext++;

            _package_.ImagePart imagePart = _mainDocumentPart.AddImagePart(_package_.ImagePartType.Jpeg, idRelationship);

            using (FileStream stream = new FileStream(@"C:\Users\Ludwig\Desktop\image1.jpeg", FileMode.Open))
                imagePart.FeedData(stream);

            _word_.Paragraph paragraph = new _word_.Paragraph();
            _word_.Run run = paragraph.AppendChild(new _word_.Run());

            _word_.Drawing drawing = run.AppendChild(new _word_.Drawing());

            _drawing_word_.Anchor anchor = drawing.AppendChild(new _drawing_word_.Anchor()
            {
                AllowOverlap = false,
                BehindDoc = true,
                DistanceFromTop = 0,
                DistanceFromBottom = 0,
                DistanceFromLeft = 0,
                DistanceFromRight = 0,
                LayoutInCell = false,
                Locked = true,
                RelativeHeight = 1,
                SimplePos = false
            });

            anchor.AppendChild(new _drawing_word_.SimplePosition()
            {
                X = 0,
                Y = 0
            });
            _drawing_word_.HorizontalPosition horizontalPosition = anchor.AppendChild(new _drawing_word_.HorizontalPosition() { RelativeFrom = _drawing_word_.HorizontalRelativePositionValues.Page });
            horizontalPosition.AppendChild(new _drawing_word_.HorizontalAlignment() { Text = alignmentString });
            _drawing_word_.VerticalPosition verticalPosition = anchor.AppendChild(new _drawing_word_.VerticalPosition() { RelativeFrom = _drawing_word_.VerticalRelativePositionValues.Page });
            verticalPosition.AppendChild(new _drawing_word_.PositionOffset() { Text = Convert.ToString(positionY) });
            anchor.AppendChild(new _drawing_word_.Extent()
            {
                Cx = imageWidth,
                Cy = imageHeight
            });
            anchor.AppendChild(new _drawing_word_.EffectExtent()
            {
                LeftEdge = 0,
                TopEdge = 0,
                RightEdge = 0,
                BottomEdge = 0
            });
            anchor.AppendChild(new _drawing_word_.WrapNone());
            anchor.AppendChild(new _drawing_word_.DocProperties()
            {
                Id = idDocElement,
                Name = "Picture " + idPicture
            });
            _drawing_word_.NonVisualGraphicFrameDrawingProperties nvgfdProperties = anchor.AppendChild(new _drawing_word_.NonVisualGraphicFrameDrawingProperties());
            nvgfdProperties.AppendChild(new _drawing_.GraphicFrameLocks() { NoChangeAspect = true });

            _drawing_.Graphic graphic = anchor.AppendChild(new _drawing_.Graphic());
            _drawing_.GraphicData graphicData = graphic.AppendChild(new _drawing_.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" });

            _drawing_pictures_.Picture picture = graphicData.AppendChild(new _drawing_pictures_.Picture());

            _drawing_pictures_.NonVisualPictureProperties nvpProperties = picture.AppendChild(new _drawing_pictures_.NonVisualPictureProperties());
            nvpProperties.AppendChild(new _drawing_pictures_.NonVisualDrawingProperties()
            {
                Id = idPicture,
                Name = "Picture " + idPicture
            });
            nvpProperties.AppendChild(new _drawing_pictures_.NonVisualPictureDrawingProperties());

            _drawing_pictures_.BlipFill blipFill = picture.AppendChild(new _drawing_pictures_.BlipFill());
            _drawing_.Blip blip = blipFill.AppendChild(new _drawing_.Blip() { Embed = idRelationship });
            _drawing_.Stretch stretch = blipFill.AppendChild(new _drawing_.Stretch());
            stretch.AppendChild(new _drawing_.FillRectangle());

            _drawing_pictures_.ShapeProperties shapeProperties = picture.AppendChild(new _drawing_pictures_.ShapeProperties());

            _drawing_.Transform2D transform2D = shapeProperties.AppendChild(new _drawing_.Transform2D());
            transform2D.AppendChild(new _drawing_.Extents()
            {
                Cx = imageWidth,
                Cy = imageHeight
            });

            _drawing_.PresetGeometry presetGeometry = shapeProperties.AppendChild(new _drawing_.PresetGeometry() { Preset = _drawing_.ShapeTypeValues.Rectangle });

            // Add element to section.

            _sections[indexSection].Elements.Add(paragraph);
            return _sections[indexSection].ElementsIndexCurrent++;
        }

        public bool CloseAndWriteToFile(string filename)
        {
            for (int i = 0; i < _sections.Count; ++i)
            {
                _sections[i].SectionProperties.AppendChild(new _word_.TitlePage());

                // Add all section elements to document body.

                foreach (_ooxml_.OpenXmlCompositeElement element in _sections[i].Elements)
                    _body.AppendChild(element);

                // Add section properties to document body.

                if (i < (_sections.Count - 1))
                {
                    _word_.Paragraph paragraph = _body.AppendChild(new _word_.Paragraph());
                    _word_.ParagraphProperties paragraphProperties = paragraph.AppendChild(new _word_.ParagraphProperties());
                    paragraphProperties.AppendChild(_sections[i].SectionProperties);
                }

                else
                    _body.AppendChild(_sections[i].SectionProperties);
            }

            bool hasError = false;

            _valid_.OpenXmlValidator validator = new _valid_.OpenXmlValidator(_ooxml_.FileFormatVersions.Office2010);

            foreach (_valid_.ValidationErrorInfo errorInfo in validator.Validate(_document))
            {
                hasError = true;

                Debug.WriteLine("*****");
                Debug.WriteLine(errorInfo.Description);
                Debug.WriteLine("-----");
                Debug.WriteLine(errorInfo.Path.XPath);
                Debug.WriteLine("*****");
            }

            if (hasError)
                return false;

            _document.Close();

            try
            {
                File.WriteAllBytes(filename, _documentStream.ToArray());

                return true;
            }

            catch (Exception e)
            {
                Debug.WriteLine(e.Message);

                return false;
            }
        }
    }
}
