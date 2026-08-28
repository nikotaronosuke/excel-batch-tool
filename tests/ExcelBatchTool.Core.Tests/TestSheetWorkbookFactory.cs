using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Charts = DocumentFormat.OpenXml.Drawing.Charts;

namespace ExcelBatchTool.Core.Tests;

/// <summary>セルに割り当てる書式(Workbook ごとに定義順で index が決まる)。</summary>
internal sealed record TestStyle
{
    public bool Bold { get; init; }

    /// <summary>塗りつぶしの ARGB(例 "FFFF0000")。</summary>
    public string? FillArgb { get; init; }

    public bool ThinBorder { get; init; }

    /// <summary>"center" など。</summary>
    public string? HorizontalAlignment { get; init; }

    /// <summary>ユーザー定義の表示形式(例 "0.00")。</summary>
    public string? NumberFormatCode { get; init; }

    /// <summary>組み込みの表示形式 ID(例 14 = 日付)。</summary>
    public uint? BuiltinNumberFormatId { get; init; }
}

/// <summary>書式を指定したセル値。<see cref="StyleId"/> は <see cref="TestStyle"/> 配列の添字。</summary>
internal sealed record Styled(object? Value, int StyleId);

internal sealed record TestRowProperty(uint RowIndex, double? Height = null, bool Hidden = false);

internal sealed record TestColumnProperty(uint Min, uint Max, double? Width = null, bool Hidden = false);

/// <summary>集約テスト用のシート定義。</summary>
internal sealed class TestAggregationSheetSpec
{
    public required string Name { get; init; }

    public bool IsHidden { get; init; }

    /// <summary>セル値。string / int / double / bool / DateTime / Styled / null。</summary>
    public object?[][] Rows { get; init; } = [];

    public TestRowProperty[] RowProperties { get; init; } = [];

    public TestColumnProperty[] ColumnProperties { get; init; } = [];

    public string[] Merges { get; init; } = [];

    public bool FreezeTopRow { get; init; }

    public bool AddProtection { get; init; }

    public bool AddFormula { get; init; }

    public bool AddChart { get; init; }

    public bool AddImage { get; init; }

    public bool AddTable { get; init; }

    public bool AddConditionalFormatting { get; init; }

    public bool AddDataValidation { get; init; }

    public bool AddHyperlink { get; init; }

    public bool AddAutoFilter { get; init; }

    public bool AddComment { get; init; }

    /// <summary>文字単位で書式を持つ共有文字列(リッチテキスト)を A1 に置く。</summary>
    public bool AddRichTextCell { get; init; }
}

/// <summary>Sheet 集約テスト用の .xlsx を架空データのみで生成する。</summary>
internal static class TestSheetWorkbookFactory
{
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    public static void Create(
        string path,
        IReadOnlyList<TestAggregationSheetSpec> sheets,
        IReadOnlyList<TestStyle>? styles = null,
        bool date1904 = false,
        bool addMacro = false,
        bool addExternalLink = false)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        if (date1904)
        {
            workbookPart.Workbook.AppendChild(new WorkbookProperties { Date1904 = true });
        }

        workbookPart.Workbook.AppendChild(new Sheets());

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = BuildStylesheet(styles ?? []);
        stylesPart.Stylesheet.Save();

        var sharedStrings = new List<SharedStringItem>();

        uint sheetId = 1;
        foreach (var spec in sheets)
        {
            AddSheet(workbookPart, spec, sheetId++, sharedStrings, date1904);
        }

        if (sharedStrings.Count > 0)
        {
            var part = workbookPart.AddNewPart<SharedStringTablePart>();
            part.SharedStringTable = new SharedStringTable(sharedStrings.Select(item => item.CloneNode(true)))
            {
                Count = (uint)sharedStrings.Count,
                UniqueCount = (uint)sharedStrings.Count,
            };
            part.SharedStringTable.Save();
        }

        if (addMacro)
        {
            var vbaPart = workbookPart.AddNewPart<VbaProjectPart>();
            using var vbaStream = new MemoryStream("架空の VBA プロジェクト"u8.ToArray());
            vbaPart.FeedData(vbaStream);
        }

        if (addExternalLink)
        {
            var externalPart = workbookPart.AddNewPart<ExternalWorkbookPart>();
            var relationship = externalPart.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
                new Uri("fictional-external.xlsx", UriKind.Relative));
            externalPart.ExternalLink = new ExternalLink(new ExternalBook { Id = relationship.Id });
            workbookPart.Workbook.AppendChild(new ExternalReferences(
                new ExternalReference { Id = workbookPart.GetIdOfPart(externalPart) }));
        }
    }

    private static void AddSheet(
        WorkbookPart workbookPart,
        TestAggregationSheetSpec spec,
        uint sheetId,
        List<SharedStringItem> sharedStrings,
        bool date1904)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        var rowProperties = spec.RowProperties.ToDictionary(property => property.RowIndex);
        var sheetData = new SheetData();
        var maxColumn = 1;

        uint rowIndex = 0;
        foreach (var values in spec.Rows)
        {
            rowIndex++;
            var row = new Row { RowIndex = rowIndex };

            if (rowProperties.TryGetValue(rowIndex, out var property))
            {
                if (property.Height is { } height)
                {
                    row.Height = height;
                    row.CustomHeight = true;
                }

                if (property.Hidden)
                {
                    row.Hidden = true;
                }
            }

            for (var i = 0; i < values.Length; i++)
            {
                if (values[i] is null)
                {
                    continue;
                }

                maxColumn = Math.Max(maxColumn, i + 1);
                row.Append(BuildCell(Reference(i + 1, rowIndex), values[i]!, sharedStrings, date1904));
            }

            sheetData.Append(row);
        }

        // 値が無く行設定だけの行も作れるようにする。
        foreach (var property in spec.RowProperties.Where(p => p.RowIndex > rowIndex))
        {
            var row = new Row { RowIndex = property.RowIndex };
            if (property.Height is { } height)
            {
                row.Height = height;
                row.CustomHeight = true;
            }

            if (property.Hidden)
            {
                row.Hidden = true;
            }

            sheetData.Append(row);
            rowIndex = Math.Max(rowIndex, property.RowIndex);
        }

        if (spec.AddFormula)
        {
            rowIndex++;
            sheetData.Append(new Row(new Cell
            {
                CellReference = Reference(1, rowIndex),
                CellFormula = new CellFormula("1+1"),
                CellValue = new CellValue("2"),
            })
            { RowIndex = rowIndex });
        }

        if (spec.AddRichTextCell)
        {
            var richIndex = sharedStrings.Count;
            sharedStrings.Add(new SharedStringItem(
                new Run(new RunProperties(new Bold()), new Text("太字")),
                new Run(new Text("通常"))));

            rowIndex++;
            sheetData.Append(new Row(new Cell
            {
                CellReference = Reference(1, rowIndex),
                DataType = CellValues.SharedString,
                CellValue = new CellValue(richIndex.ToString(CultureInfo.InvariantCulture)),
            })
            { RowIndex = rowIndex });
        }

        var lastRow = Math.Max(1u, rowIndex);
        var children = new List<OpenXmlElement>
        {
            new SheetDimension { Reference = $"A1:{Letters(maxColumn)}{lastRow}" },
            BuildSheetViews(spec),
        };

        if (spec.ColumnProperties.Length > 0)
        {
            var columns = new Columns();
            foreach (var property in spec.ColumnProperties)
            {
                var column = new Column { Min = property.Min, Max = property.Max };
                if (property.Width is { } width)
                {
                    column.Width = width;
                    column.CustomWidth = true;
                }

                if (property.Hidden)
                {
                    column.Hidden = true;
                }

                columns.Append(column);
            }

            children.Add(columns);
        }

        children.Add(sheetData);

        if (spec.AddProtection)
        {
            children.Add(new SheetProtection { Sheet = true, Objects = true, Scenarios = true });
        }

        if (spec.AddAutoFilter)
        {
            children.Add(new AutoFilter { Reference = $"A1:{Letters(maxColumn)}{lastRow}" });
        }

        if (spec.Merges.Length > 0)
        {
            children.Add(new MergeCells(spec.Merges.Select(reference => new MergeCell { Reference = reference }))
            {
                Count = (uint)spec.Merges.Length,
            });
        }

        if (spec.AddConditionalFormatting)
        {
            children.Add(new ConditionalFormatting(
                new ConditionalFormattingRule(new Formula("1"))
                {
                    Type = ConditionalFormatValues.Expression,
                    Priority = 1,
                    FormatId = 0U,
                })
            { SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A5" } });
        }

        if (spec.AddDataValidation)
        {
            children.Add(new DataValidations(
                new DataValidation(new Formula1("1"), new Formula2("100"))
                {
                    Type = DataValidationValues.Whole,
                    Operator = DataValidationOperatorValues.Between,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A5" },
                })
            { Count = 1U });
        }

        var worksheet = new Worksheet();
        foreach (var child in children)
        {
            worksheet.Append(child);
        }

        worksheetPart.Worksheet = worksheet;

        if (spec.AddHyperlink)
        {
            var relationship = worksheetPart.AddHyperlinkRelationship(
                new Uri("https://example.invalid/", UriKind.Absolute), isExternal: true);
            worksheet.Append(new Hyperlinks(new Hyperlink { Reference = "A1", Id = relationship.Id }));
        }

        if (spec.AddComment)
        {
            var commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>();
            commentsPart.Comments = new Comments(
                new Authors(new Author("架空作成者")),
                new CommentList(new Comment(new CommentText(new Text("架空コメント")))
                {
                    Reference = "A1",
                    AuthorId = 0U,
                }));
            commentsPart.Comments.Save();
        }

        if (spec.AddChart || spec.AddImage)
        {
            var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();

            if (spec.AddChart)
            {
                var chartPart = drawingsPart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = new Charts.ChartSpace(new Charts.Chart(new Charts.PlotArea()));
            }

            if (spec.AddImage)
            {
                var imagePart = drawingsPart.AddImagePart("image/png");
                using var stream = new MemoryStream(TinyPng);
                imagePart.FeedData(stream);
            }

            worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
        }

        if (spec.AddTable)
        {
            var tablePart = worksheetPart.AddNewPart<TableDefinitionPart>();
            tablePart.Table = new Table(
                new TableColumns(new TableColumn { Id = 1U, Name = "列1" }) { Count = 1U })
            {
                Id = 1U,
                Name = "架空テーブル",
                DisplayName = "架空テーブル",
                Reference = "A1:A2",
                TotalsRowShown = false,
            };
            tablePart.Table.Save();
            worksheet.Append(new TableParts(new TablePart { Id = worksheetPart.GetIdOfPart(tablePart) }) { Count = 1U });
        }

        var sheet = new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = spec.Name,
        };

        if (spec.IsHidden)
        {
            sheet.State = SheetStateValues.Hidden;
        }

        workbookPart.Workbook!.Sheets!.Append(sheet);
    }

    private static SheetViews BuildSheetViews(TestAggregationSheetSpec spec)
    {
        var view = new SheetView { WorkbookViewId = 0U };
        if (spec.FreezeTopRow)
        {
            view.Append(new Pane
            {
                VerticalSplit = 1D,
                TopLeftCell = "A2",
                ActivePane = PaneValues.BottomLeft,
                State = PaneStateValues.Frozen,
            });
        }

        return new SheetViews(view);
    }

    private static Cell BuildCell(string reference, object value, List<SharedStringItem> sharedStrings, bool date1904)
    {
        var styleId = 0U;
        if (value is Styled styled)
        {
            styleId = (uint)(styled.StyleId + 1); // 0 は既定書式
            value = styled.Value ?? string.Empty;
        }

        var cell = value switch
        {
            string text => new Cell
            {
                CellReference = reference,
                DataType = CellValues.SharedString,
                CellValue = new CellValue(AddSharedString(sharedStrings, text).ToString(CultureInfo.InvariantCulture)),
            },

            bool flag => new Cell
            {
                CellReference = reference,
                DataType = CellValues.Boolean,
                CellValue = new CellValue(flag ? "1" : "0"),
            },

            DateTime date => new Cell
            {
                CellReference = reference,
                CellValue = new CellValue(ToSerial(date, date1904).ToString(CultureInfo.InvariantCulture)),
            },

            _ => new Cell
            {
                CellReference = reference,
                CellValue = new CellValue(Convert.ToDouble(value).ToString(CultureInfo.InvariantCulture)),
            },
        };

        if (styleId != 0)
        {
            cell.StyleIndex = styleId;
        }

        return cell;
    }

    private static double ToSerial(DateTime value, bool date1904)
        => date1904 ? value.ToOADate() - 1462 : value.ToOADate();

    private static int AddSharedString(List<SharedStringItem> sharedStrings, string text)
    {
        for (var i = 0; i < sharedStrings.Count; i++)
        {
            if (sharedStrings[i].Descendants<Run>().Any())
            {
                continue; // リッチテキストは共有しない。
            }

            if (string.Equals(sharedStrings[i].InnerText, text, StringComparison.Ordinal))
            {
                return i;
            }
        }

        sharedStrings.Add(new SharedStringItem(new Text(text)));
        return sharedStrings.Count - 1;
    }

    private static Stylesheet BuildStylesheet(IReadOnlyList<TestStyle> styles)
    {
        var fonts = new Fonts(new Font(new FontSize { Val = 11D }));
        var fills = new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
        var borders = new Borders(new Border());
        var numberingFormats = new NumberingFormats();
        var cellFormats = new CellFormats(new CellFormat
        {
            NumberFormatId = 0U,
            FontId = 0U,
            FillId = 0U,
            BorderId = 0U,
            FormatId = 0U,
        });

        uint nextNumberFormatId = 164;

        foreach (var style in styles)
        {
            var format = new CellFormat { FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U, NumberFormatId = 0U };

            if (style.Bold)
            {
                fonts.Append(new Font(new Bold(), new FontSize { Val = 11D }));
                format.FontId = (uint)fonts.Count() - 1;
                format.ApplyFont = true;
            }

            if (style.FillArgb is { } argb)
            {
                fills.Append(new Fill(new PatternFill(new ForegroundColor { Rgb = argb })
                {
                    PatternType = PatternValues.Solid,
                }));
                format.FillId = (uint)fills.Count() - 1;
                format.ApplyFill = true;
            }

            if (style.ThinBorder)
            {
                borders.Append(new Border(
                    new LeftBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                    new RightBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                    new TopBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                    new BottomBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                    new DiagonalBorder()));
                format.BorderId = (uint)borders.Count() - 1;
                format.ApplyBorder = true;
            }

            if (style.HorizontalAlignment is { } alignment)
            {
                format.Append(new Alignment { Horizontal = new EnumValue<HorizontalAlignmentValues>(
                    new HorizontalAlignmentValues(alignment)) });
                format.ApplyAlignment = true;
            }

            if (style.NumberFormatCode is { } code)
            {
                var id = nextNumberFormatId++;
                numberingFormats.Append(new NumberingFormat { NumberFormatId = id, FormatCode = code });
                format.NumberFormatId = id;
                format.ApplyNumberFormat = true;
            }
            else if (style.BuiltinNumberFormatId is { } builtinId)
            {
                format.NumberFormatId = builtinId;
                format.ApplyNumberFormat = true;
            }

            cellFormats.Append(format);
        }

        fonts.Count = (uint)fonts.Count();
        fills.Count = (uint)fills.Count();
        borders.Count = (uint)borders.Count();
        cellFormats.Count = (uint)cellFormats.Count();

        var stylesheet = new Stylesheet();
        if (numberingFormats.Any())
        {
            numberingFormats.Count = (uint)numberingFormats.Count();
            stylesheet.Append(numberingFormats);
        }

        stylesheet.Append(fonts);
        stylesheet.Append(fills);
        stylesheet.Append(borders);
        stylesheet.Append(new CellStyleFormats(
            new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U })
        { Count = 1U });
        stylesheet.Append(cellFormats);
        stylesheet.Append(new CellStyles(new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U }) { Count = 1U });

        return stylesheet;
    }

    private static string Reference(int column, uint row) => $"{Letters(column)}{row}";

    private static string Letters(int columnIndex) => CellRangeParser.ColumnIndexToLetters(columnIndex);
}
