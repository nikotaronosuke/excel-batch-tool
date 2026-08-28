using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Charts = DocumentFormat.OpenXml.Drawing.Charts;

namespace ExcelBatchTool.Core.Tests;

/// <summary>一括変更テスト用のセル。<see cref="StyleId"/> は <see cref="MutationTestStyle"/> 配列の添字 + 1。</summary>
internal sealed record MutationTestCell(string Reference, object? Value, int? StyleId = null);

/// <summary>一括変更テスト用の書式(表示形式だけを変える)。</summary>
internal sealed record MutationTestStyle(uint NumberFormatId = 0, string? FormatCode = null);

/// <summary>一括変更テスト用のシート定義。</summary>
internal sealed class MutationTestSheet
{
    public required string Name { get; init; }

    public MutationTestCell[] Cells { get; init; } = [];

    public string[] Merges { get; init; } = [];

    public bool AddProtection { get; init; }

    /// <summary>入力規則を付ける適用範囲(sqref)。</summary>
    public string? DataValidationSqref { get; init; }

    /// <summary>Office 2010 以降の拡張入力規則を付ける適用範囲。</summary>
    public string? X14ValidationSqref { get; init; }

    /// <summary>ハイパーリンクを付けるセル。</summary>
    public string? HyperlinkReference { get; init; }

    /// <summary>数式セルを 1 つ追加する。</summary>
    public string? FormulaCell { get; init; }

    /// <summary>リッチテキスト(文字ごとの書式)を持つセル。</summary>
    public string? RichTextCell { get; init; }

    /// <summary>セル値のメタデータ参照(vm)を付けるセル。</summary>
    public string? MetadataCell { get; init; }

    public bool AddChart { get; init; }

    public bool AddImage { get; init; }

    public bool AddTable { get; init; }

    public bool AddPivotTable { get; init; }

    public bool AddConditionalFormatting { get; init; }

    /// <summary>Excel の形式として不正な内容(min の無い col)を入れる。</summary>
    public bool AddSchemaError { get; init; }
}

/// <summary>一括変更テスト用の .xlsx を架空データのみで生成する。</summary>
internal static class TestMutationWorkbookFactory
{
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    public static void Create(
        string path,
        IReadOnlyList<MutationTestSheet> sheets,
        IReadOnlyList<MutationTestStyle>? styles = null,
        bool addDocumentProperties = true,
        bool addDefinedName = true,
        bool addExternalLink = false)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        workbookPart.Workbook.AppendChild(new Sheets());

        AddTheme(workbookPart);

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = BuildStylesheet(styles ?? []);
        stylesPart.Stylesheet.Save();

        var sharedStrings = new List<SharedStringItem>();
        var pivotCacheIds = new List<string>();

        uint sheetId = 1;
        foreach (var sheet in sheets)
        {
            AddSheet(workbookPart, sheet, sheetId++, sharedStrings, pivotCacheIds);
        }

        // 共有文字列は必ず作る(セル値が数値だけでも、Part が消えないことを確かめたいため)。
        sharedStrings.Add(new SharedStringItem(new Text("架空の予備文字列")));

        var sharedPart = workbookPart.AddNewPart<SharedStringTablePart>();
        sharedPart.SharedStringTable = new SharedStringTable(sharedStrings.Select(item => item.CloneNode(true)))
        {
            Count = (uint)sharedStrings.Count,
            UniqueCount = (uint)sharedStrings.Count,
        };
        sharedPart.SharedStringTable.Save();

        // CT_Workbook の子要素順: sheets → externalReferences → definedNames → pivotCaches。
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

        if (addDefinedName && sheets.Count > 0)
        {
            workbookPart.Workbook.AppendChild(new DefinedNames(
                new DefinedName($"'{sheets[0].Name}'!$A$1") { Name = "架空の名前" }));
        }

        if (pivotCacheIds.Count > 0)
        {
            uint cacheId = 1;
            workbookPart.Workbook.AppendChild(new PivotCaches(
                pivotCacheIds.Select(id => new PivotCache { CacheId = cacheId++, Id = id })));
        }

        workbookPart.Workbook.Save();

        if (addDocumentProperties)
        {
            document.PackageProperties.Title = "架空のブック";
            document.PackageProperties.Creator = "架空作成者";
        }
    }

    private static void AddSheet(
        WorkbookPart workbookPart,
        MutationTestSheet spec,
        uint sheetId,
        List<SharedStringItem> sharedStrings,
        List<string> pivotCacheIds)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        // 行ごとにまとめてから、行番号・列順で並べて書く。
        var rows = new SortedDictionary<uint, SortedDictionary<int, Cell>>();

        foreach (var cell in spec.Cells)
        {
            Place(rows, cell.Reference, BuildCell(cell, sharedStrings));
        }

        if (spec.FormulaCell is { } formulaCell)
        {
            Place(rows, formulaCell, new Cell
            {
                CellReference = formulaCell,
                CellFormula = new CellFormula("1+1"),
                CellValue = new CellValue("2"),
            });
        }

        if (spec.RichTextCell is { } richCell)
        {
            var index = sharedStrings.Count;
            sharedStrings.Add(new SharedStringItem(
                new Run(new RunProperties(new Bold()), new Text("太字")),
                new Run(new Text("通常"))));

            Place(rows, richCell, new Cell
            {
                CellReference = richCell,
                DataType = CellValues.SharedString,
                CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture)),
            });
        }

        if (spec.MetadataCell is { } metadataCell)
        {
            AddCellMetadata(workbookPart);
            Place(rows, metadataCell, new Cell
            {
                CellReference = metadataCell,
                ValueMetaIndex = 1U,
                CellValue = new CellValue("1"),
            });
        }

        var sheetData = new SheetData();
        foreach (var (rowIndex, cells) in rows)
        {
            var row = new Row { RowIndex = rowIndex };
            foreach (var cell in cells.Values)
            {
                row.Append(cell);
            }

            sheetData.Append(row);
        }

        // CT_Worksheet の子要素順で組み立てる。
        var children = new List<OpenXmlElement>
        {
            new SheetViews(new SheetView { WorkbookViewId = 0U }),
        };

        if (spec.AddSchemaError)
        {
            // col には min が必須。開けるが Excel の形式としては不正な状態を作る。
            children.Add(new Columns(new Column { Max = 3U, Width = 12D, CustomWidth = true }));
        }

        children.Add(sheetData);

        if (spec.AddProtection)
        {
            children.Add(new SheetProtection { Sheet = true, Objects = true, Scenarios = true });
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
            children.Add(new ConditionalFormatting(new ConditionalFormattingRule
            {
                Type = ConditionalFormatValues.DuplicateValues,
                FormatId = 0U,
                Priority = 1,
            })
            { SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A20" } });
        }

        if (spec.DataValidationSqref is { } sqref)
        {
            children.Add(new DataValidations(
                new DataValidation(new Formula1("\"はい,いいえ\""))
                {
                    Type = DataValidationValues.List,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = sqref },
                })
            { Count = 1U });
        }

        if (spec.HyperlinkReference is { } linkReference)
        {
            var relationship = worksheetPart.AddHyperlinkRelationship(
                new Uri("https://example.invalid/", UriKind.Absolute), isExternal: true);
            children.Add(new Hyperlinks(new Hyperlink
            {
                Reference = linkReference,
                Id = relationship.Id,
            }));
        }

        children.Add(new PageMargins
        {
            Left = 0.7D,
            Right = 0.7D,
            Top = 0.75D,
            Bottom = 0.75D,
            Header = 0.3D,
            Footer = 0.3D,
        });

        var worksheet = new Worksheet();
        foreach (var child in children)
        {
            worksheet.Append(child);
        }

        worksheetPart.Worksheet = worksheet;

        if (spec.AddChart || spec.AddImage)
        {
            var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();

            if (spec.AddChart)
            {
                var chartPart = drawingsPart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = new Charts.ChartSpace(
                    new Charts.Chart(
                        new Charts.PlotArea(
                            new Charts.Layout(),
                            new Charts.BarChart(
                                new Charts.BarDirection { Val = Charts.BarDirectionValues.Column },
                                new Charts.BarGrouping { Val = Charts.BarGroupingValues.Clustered },
                                new Charts.AxisId { Val = 1U },
                                new Charts.AxisId { Val = 2U }),
                            new Charts.CategoryAxis(
                                new Charts.AxisId { Val = 1U },
                                new Charts.Scaling(new Charts.Orientation
                                {
                                    Val = Charts.OrientationValues.MinMax,
                                }),
                                new Charts.Delete { Val = false },
                                new Charts.AxisPosition { Val = Charts.AxisPositionValues.Bottom },
                                new Charts.CrossingAxis { Val = 2U }),
                            new Charts.ValueAxis(
                                new Charts.AxisId { Val = 2U },
                                new Charts.Scaling(new Charts.Orientation
                                {
                                    Val = Charts.OrientationValues.MinMax,
                                }),
                                new Charts.Delete { Val = false },
                                new Charts.AxisPosition { Val = Charts.AxisPositionValues.Left },
                                new Charts.CrossingAxis { Val = 1U }))));
                chartPart.ChartSpace.Save();
            }

            if (spec.AddImage)
            {
                var imagePart = drawingsPart.AddImagePart("image/png");
                using var stream = new MemoryStream(TinyPng);
                imagePart.FeedData(stream);
            }

            drawingsPart.WorksheetDrawing.Save();
            worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
        }

        if (spec.AddTable)
        {
            var tablePart = worksheetPart.AddNewPart<TableDefinitionPart>();
            tablePart.Table = new Table(
                new TableColumns(
                    new TableColumn { Id = 1U, Name = "項目" },
                    new TableColumn { Id = 2U, Name = "値" })
                { Count = 2U })
            {
                Id = 1U,
                Name = "架空テーブル",
                DisplayName = "架空テーブル",
                Reference = "D1:E2",
                TotalsRowShown = false,
            };
            tablePart.Table.Save();

            worksheet.Append(new TableParts(new TablePart { Id = worksheetPart.GetIdOfPart(tablePart) })
            {
                Count = 1U,
            });
        }

        if (spec.AddPivotTable)
        {
            pivotCacheIds.Add(
                AddPivotTable(workbookPart, worksheetPart, spec.Name, (uint)pivotCacheIds.Count + 1));
        }

        if (spec.X14ValidationSqref is { } x14Sqref)
        {
            var container = new DocumentFormat.OpenXml.Office2010.Excel.DataValidations(
                new DocumentFormat.OpenXml.Office2010.Excel.DataValidation(
                    new DocumentFormat.OpenXml.Office2010.Excel.DataValidationForumla1(
                        new DocumentFormat.OpenXml.Office.Excel.Formula("\"はい,いいえ\"")),
                    new DocumentFormat.OpenXml.Office.Excel.ReferenceSequence(x14Sqref))
                {
                    Type = DataValidationValues.List,
                })
            { Count = 1U };

            var extension = new WorksheetExtension(container)
            {
                Uri = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}",
            };
            extension.AddNamespaceDeclaration(
                "x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            worksheet.Append(new WorksheetExtensionList(extension));
        }

        workbookPart.Workbook!.Sheets!.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = spec.Name,
        });
    }

    /// <summary>セル値に紐づくメタデータ(vm で参照される)を作る。</summary>
    private static void AddCellMetadata(WorkbookPart workbookPart)
    {
        if (workbookPart.GetPartsOfType<CellMetadataPart>().Any())
        {
            return;
        }

        var part = workbookPart.AddNewPart<CellMetadataPart>();
        part.Metadata = new Metadata(
            new MetadataTypes(new MetadataType
            {
                Name = "架空メタデータ",
                MinSupportedVersion = 120000U,
                Copy = true,
                PasteAll = true,
                PasteValues = true,
                Merge = true,
                SplitFirst = true,
                RowColumnShift = true,
                ClearFormats = true,
                ClearComments = true,
                Assign = true,
                Coerce = true,
            })
            { Count = 1U },
            new ValueMetadata(
                new MetadataBlock(new MetadataRecord { TypeIndex = 1U, Val = 0U }))
            { Count = 1U });

        part.Metadata.Save();
    }

    /// <summary>ピボットテーブル一式(キャッシュ定義・レコード・定義)を作る。</summary>
    private static string AddPivotTable(
        WorkbookPart workbookPart,
        WorksheetPart worksheetPart,
        string sheetName,
        uint cacheId)
    {
        var cachePart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>();
        var recordsPart = cachePart.AddNewPart<PivotTableCacheRecordsPart>();
        recordsPart.PivotCacheRecords = new PivotCacheRecords { Count = 0U };
        recordsPart.PivotCacheRecords.Save();

        cachePart.PivotCacheDefinition = new PivotCacheDefinition(
            new CacheSource(new WorksheetSource { Reference = "A1:B2", Sheet = sheetName })
            {
                Type = SourceValues.Worksheet,
            },
            new CacheFields(
                new CacheField(new SharedItems()) { Name = "項目", NumberFormatId = 0U },
                new CacheField(new SharedItems()) { Name = "値", NumberFormatId = 0U })
            { Count = 2U })
        {
            Id = cachePart.GetIdOfPart(recordsPart),
            RecordCount = 0U,
        };
        cachePart.PivotCacheDefinition.Save();

        var pivotPart = worksheetPart.AddNewPart<PivotTablePart>();
        pivotPart.AddPart(cachePart);
        pivotPart.PivotTableDefinition = new PivotTableDefinition(
            new Location
            {
                Reference = "H1:I3",
                FirstHeaderRow = 1U,
                FirstDataRow = 1U,
                FirstDataColumn = 1U,
            },
            new PivotFields(
                new PivotField { ShowAll = false },
                new PivotField { ShowAll = false })
            { Count = 2U })
        {
            Name = "架空ピボット",
            CacheId = cacheId,
            DataCaption = "値",
        };
        pivotPart.PivotTableDefinition.Save();

        return workbookPart.GetIdOfPart(cachePart);
    }

    private static void Place(
        SortedDictionary<uint, SortedDictionary<int, Cell>> rows, string reference, Cell cell)
    {
        if (!CellRangeParser.TryParseCell(reference, out var column, out var row))
        {
            throw new ArgumentException($"セル参照を解釈できません: {reference}", nameof(reference));
        }

        if (!rows.TryGetValue((uint)row, out var cells))
        {
            cells = [];
            rows[(uint)row] = cells;
        }

        cells[column] = cell;
    }

    private static Cell BuildCell(MutationTestCell spec, List<SharedStringItem> sharedStrings)
    {
        var cell = new Cell { CellReference = spec.Reference };

        if (spec.StyleId is { } styleId)
        {
            cell.StyleIndex = (uint)styleId;
        }

        switch (spec.Value)
        {
            case null:
                break;

            case string text:
                var index = sharedStrings.Count;
                sharedStrings.Add(new SharedStringItem(new Text(text)));
                cell.DataType = CellValues.SharedString;
                cell.CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture));
                break;

            case bool flag:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(flag ? "1" : "0");
                break;

            default:
                cell.CellValue = new CellValue(
                    Convert.ToDouble(spec.Value, CultureInfo.InvariantCulture)
                        .ToString(CultureInfo.InvariantCulture));
                break;
        }

        return cell;
    }

    private static void AddTheme(WorkbookPart workbookPart)
    {
        var themePart = workbookPart.AddNewPart<ThemePart>();
        themePart.Theme = new DocumentFormat.OpenXml.Drawing.Theme(
            new DocumentFormat.OpenXml.Drawing.ThemeElements(
                new DocumentFormat.OpenXml.Drawing.ColorScheme(
                    new DocumentFormat.OpenXml.Drawing.Dark1Color(
                        new DocumentFormat.OpenXml.Drawing.SystemColor
                        {
                            Val = DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowText,
                        }),
                    new DocumentFormat.OpenXml.Drawing.Light1Color(
                        new DocumentFormat.OpenXml.Drawing.SystemColor
                        {
                            Val = DocumentFormat.OpenXml.Drawing.SystemColorValues.Window,
                        }),
                    new DocumentFormat.OpenXml.Drawing.Dark2Color(
                        new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "1F497D" }),
                    new DocumentFormat.OpenXml.Drawing.Light2Color(
                        new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "EEECE1" }),
                    new DocumentFormat.OpenXml.Drawing.Accent1Color(
                        new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "4F81BD" }),
                    new DocumentFormat.OpenXml.Drawing.Accent2Color(
                        new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "C0504D" }),
                    new DocumentFormat.OpenXml.Drawing.Accent3Color(
                        new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "9BBB59" }),
                    new DocumentFormat.OpenXml.Drawing.Accent4Color(
                        new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "8064A2" }),
                    new DocumentFormat.OpenXml.Drawing.Accent5Color(
                        new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "4BACC6" }),
                    new DocumentFormat.OpenXml.Drawing.Accent6Color(
                        new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "F79646" }),
                    new DocumentFormat.OpenXml.Drawing.Hyperlink(
                        new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "0000FF" }),
                    new DocumentFormat.OpenXml.Drawing.FollowedHyperlinkColor(
                        new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "800080" }))
                { Name = "架空配色" },
                new DocumentFormat.OpenXml.Drawing.FontScheme(
                    new DocumentFormat.OpenXml.Drawing.MajorFont(
                        new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "Calibri" },
                        new DocumentFormat.OpenXml.Drawing.EastAsianFont { Typeface = string.Empty },
                        new DocumentFormat.OpenXml.Drawing.ComplexScriptFont { Typeface = string.Empty }),
                    new DocumentFormat.OpenXml.Drawing.MinorFont(
                        new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "Calibri" },
                        new DocumentFormat.OpenXml.Drawing.EastAsianFont { Typeface = string.Empty },
                        new DocumentFormat.OpenXml.Drawing.ComplexScriptFont { Typeface = string.Empty }))
                { Name = "架空フォント" },
                new DocumentFormat.OpenXml.Drawing.FormatScheme(
                    new DocumentFormat.OpenXml.Drawing.FillStyleList(
                        new DocumentFormat.OpenXml.Drawing.SolidFill(
                            new DocumentFormat.OpenXml.Drawing.SchemeColor
                            {
                                Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor,
                            }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(
                            new DocumentFormat.OpenXml.Drawing.SchemeColor
                            {
                                Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor,
                            }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(
                            new DocumentFormat.OpenXml.Drawing.SchemeColor
                            {
                                Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor,
                            })),
                    new DocumentFormat.OpenXml.Drawing.LineStyleList(
                        new DocumentFormat.OpenXml.Drawing.Outline(
                            new DocumentFormat.OpenXml.Drawing.SolidFill(
                                new DocumentFormat.OpenXml.Drawing.SchemeColor
                                {
                                    Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor,
                                })),
                        new DocumentFormat.OpenXml.Drawing.Outline(
                            new DocumentFormat.OpenXml.Drawing.SolidFill(
                                new DocumentFormat.OpenXml.Drawing.SchemeColor
                                {
                                    Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor,
                                })),
                        new DocumentFormat.OpenXml.Drawing.Outline(
                            new DocumentFormat.OpenXml.Drawing.SolidFill(
                                new DocumentFormat.OpenXml.Drawing.SchemeColor
                                {
                                    Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor,
                                }))),
                    new DocumentFormat.OpenXml.Drawing.EffectStyleList(
                        new DocumentFormat.OpenXml.Drawing.EffectStyle(
                            new DocumentFormat.OpenXml.Drawing.EffectList()),
                        new DocumentFormat.OpenXml.Drawing.EffectStyle(
                            new DocumentFormat.OpenXml.Drawing.EffectList()),
                        new DocumentFormat.OpenXml.Drawing.EffectStyle(
                            new DocumentFormat.OpenXml.Drawing.EffectList())),
                    new DocumentFormat.OpenXml.Drawing.BackgroundFillStyleList(
                        new DocumentFormat.OpenXml.Drawing.SolidFill(
                            new DocumentFormat.OpenXml.Drawing.SchemeColor
                            {
                                Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor,
                            }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(
                            new DocumentFormat.OpenXml.Drawing.SchemeColor
                            {
                                Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor,
                            }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(
                            new DocumentFormat.OpenXml.Drawing.SchemeColor
                            {
                                Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor,
                            })))
                { Name = "架空書式" }))
        { Name = "架空テーマ" };

        themePart.Theme.Save();
    }

    /// <summary>先頭は既定(General)。以降は指定された表示形式を 1 つずつ持つ。</summary>
    private static Stylesheet BuildStylesheet(IReadOnlyList<MutationTestStyle> styles)
    {
        var numberingFormats = new NumberingFormats();
        var cellFormats = new CellFormats(new CellFormat
        {
            NumberFormatId = 0U,
            FontId = 0U,
            FillId = 0U,
            BorderId = 0U,
            FormatId = 0U,
        });

        foreach (var style in styles)
        {
            var id = style.NumberFormatId;
            if (style.FormatCode is { } code)
            {
                numberingFormats.Append(new NumberingFormat { NumberFormatId = id, FormatCode = code });
            }

            cellFormats.Append(new CellFormat
            {
                NumberFormatId = id,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
                ApplyNumberFormat = true,
            });
        }

        var stylesheet = new Stylesheet();
        if (numberingFormats.Any())
        {
            numberingFormats.Count = (uint)numberingFormats.Count();
            stylesheet.Append(numberingFormats);
        }

        stylesheet.Append(new Fonts(new Font(new FontSize { Val = 11D })) { Count = 1U });
        stylesheet.Append(new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2U });
        stylesheet.Append(new Borders(new Border()) { Count = 1U });
        stylesheet.Append(new CellStyleFormats(
            new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U }) { Count = 1U });

        cellFormats.Count = (uint)cellFormats.Count();
        stylesheet.Append(cellFormats);
        stylesheet.Append(new CellStyles(
            new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U }) { Count = 1U });
        stylesheet.Append(new DifferentialFormats(
            new DifferentialFormat(new Font(new Bold()))) { Count = 1U });

        return stylesheet;
    }
}
