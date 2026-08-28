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

/// <summary>
/// テスト用のハイパーリンク指定。
/// <paramref name="ExternalTarget"/> を指定すると relationship 付きの外部リンク、
/// <paramref name="Location"/> だけならブック内リンクになる。
/// </summary>
internal sealed record TestHyperlink(
    string Reference,
    string? ExternalTarget = null,
    string? Location = null,
    string? Tooltip = null,
    string? Display = null,
    bool ExternalTargetIsRelative = false,
    bool UseDanglingRelationshipId = false,
    bool UseInternalRelationship = false);

/// <summary>テスト用の入力規則指定。</summary>
internal sealed record TestDataValidation(
    string Sqref,
    string Type,
    string? Operator = null,
    string? Formula1 = null,
    string? Formula2 = null,
    bool AllowBlank = false,
    bool ShowDropDown = false,
    bool ShowInputMessage = false,
    bool ShowErrorMessage = false,
    string? ErrorStyle = null,
    string? ImeMode = null,
    string? PromptTitle = null,
    string? Prompt = null,
    string? ErrorTitle = null,
    string? Error = null);

/// <summary>集約テスト用のシート定義。</summary>
internal sealed class TestAggregationSheetSpec
{
    public required string Name { get; init; }

    public bool IsHidden { get; init; }

    /// <summary>「非常に非表示」。<see cref="IsHidden"/> より優先する。</summary>
    public bool IsVeryHidden { get; init; }

    public bool AddPageMargins { get; init; }

    public bool AddPageSetup { get; init; }

    public bool AddPrintOptions { get; init; }

    public bool AddHeaderFooter { get; init; }

    public bool AddRowBreaks { get; init; }

    public bool AddColumnBreaks { get; init; }

    /// <summary>シート見出しの色(sheetPr/tabColor)を付ける。</summary>
    public bool AddTabColor { get; init; }

    /// <summary>印刷の拡大縮小設定(sheetPr/pageSetUpPr)を付ける。</summary>
    public bool AddPageSetupProperties { get; init; }

    /// <summary>ふりがな設定(phoneticPr)を付ける。</summary>
    public bool AddPhoneticProperties { get; init; }

    /// <summary>pageSetup にプリンター設定パートへの r:id を持たせる。</summary>
    public bool AddPrinterSettings { get; init; }

    /// <summary>ヘッダー・フッターの画像(drawingHF)を付ける。</summary>
    public bool AddHeaderFooterDrawing { get; init; }

    /// <summary>ヘッダー文字列に画像コード(&amp;G)を入れる。</summary>
    public bool AddHeaderFooterImageCode { get; init; }

    /// <summary>ヘッダー・フッターの奇数偶数・先頭ページ区別を有効にする。</summary>
    public bool AddDistinctHeaderFooter { get; init; }

    /// <summary>壊れた改ページ定義(位置 0)を入れる。</summary>
    public bool AddBrokenBreak { get; init; }

    /// <summary>_xlnm.Print_Area の参照文字列(シート名込み)。null なら作らない。</summary>
    public string? PrintArea { get; init; }

    /// <summary>_xlnm.Print_Titles の参照文字列(シート名込み)。null なら作らない。</summary>
    public string? PrintTitles { get; init; }

    /// <summary>このシートに固有の、印刷以外の名前定義。</summary>
    public (string Name, string Reference)? LocalDefinedName { get; init; }

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

    /// <summary>個別に指定する入力規則。</summary>
    public TestDataValidation[] DataValidations { get; init; } = [];

    /// <summary>入力規則コンテナの属性(disablePrompts / xWindow / yWindow)を付ける。</summary>
    public bool AddDataValidationContainerAttributes { get; init; }

    /// <summary>入力規則に想定外の拡張属性を付ける。</summary>
    public bool AddUnknownDataValidationAttribute { get; init; }

    /// <summary>Office 2010 以降の拡張入力規則(x14)を extLst へ入れる。</summary>
    public bool AddX14DataValidation { get; init; }

    public bool AddHyperlink { get; init; }

    /// <summary>個別に指定するハイパーリンク。</summary>
    public TestHyperlink[] Hyperlinks { get; init; } = [];

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

        // Defined Names は sheets の後ろ。localSheetId はシートの並び順。
        var definedNames = new DefinedNames();
        for (var index = 0; index < sheets.Count; index++)
        {
            var spec = sheets[index];
            if (spec.PrintArea is { } printArea)
            {
                definedNames.Append(new DefinedName(printArea)
                {
                    Name = "_xlnm.Print_Area",
                    LocalSheetId = (uint)index,
                });
            }

            if (spec.PrintTitles is { } printTitles)
            {
                definedNames.Append(new DefinedName(printTitles)
                {
                    Name = "_xlnm.Print_Titles",
                    LocalSheetId = (uint)index,
                });
            }

            if (spec.LocalDefinedName is { } local)
            {
                definedNames.Append(new DefinedName(local.Reference)
                {
                    Name = local.Name,
                    LocalSheetId = (uint)index,
                });
            }
        }

        if (definedNames.Any())
        {
            workbookPart.Workbook.AppendChild(definedNames);
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
        var children = new List<OpenXmlElement>();

        if (spec.AddTabColor || spec.AddPageSetupProperties)
        {
            var sheetProperties = new SheetProperties();
            if (spec.AddTabColor)
            {
                sheetProperties.TabColor = new TabColor { Rgb = "FF0000FF" };
            }

            if (spec.AddPageSetupProperties)
            {
                sheetProperties.PageSetupProperties = new PageSetupProperties { FitToPage = true };
            }

            children.Add(sheetProperties);
        }

        children.Add(new SheetDimension { Reference = $"A1:{Letters(maxColumn)}{lastRow}" });
        children.Add(BuildSheetViews(spec));

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

        if (spec.AddPhoneticProperties)
        {
            children.Add(new PhoneticProperties { FontId = 0U });
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

        if (spec.AddDataValidation || spec.DataValidations.Length > 0)
        {
            var container = new DataValidations();

            if (spec.AddDataValidationContainerAttributes)
            {
                container.DisablePrompts = true;
                container.XWindow = 100U;
                container.YWindow = 200U;
            }

            if (spec.AddDataValidation)
            {
                container.Append(new DataValidation(new Formula1("1"), new Formula2("100"))
                {
                    Type = DataValidationValues.Whole,
                    Operator = DataValidationOperatorValues.Between,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A5" },
                });
            }

            foreach (var validation in spec.DataValidations)
            {
                container.Append(BuildDataValidation(validation, spec.AddUnknownDataValidationAttribute));
            }

            // count を実際と違う値にしておき、出力側で振り直されることを確かめられるようにする。
            container.Count = 99U;
            children.Add(container);
        }

        // CT_Worksheet の要素順に合わせて印刷/ページレイアウト系を追加する。
        if (spec.AddPrintOptions)
        {
            children.Add(new PrintOptions { HorizontalCentered = true });
        }

        if (spec.AddPageMargins)
        {
            children.Add(new PageMargins
            {
                Left = 0.7D,
                Right = 0.7D,
                Top = 0.75D,
                Bottom = 0.75D,
                Header = 0.3D,
                Footer = 0.3D,
            });
        }

        if (spec.AddPageSetup)
        {
            var pageSetup = new PageSetup
            {
                PaperSize = 9U,
                Orientation = OrientationValues.Landscape,
                Scale = 85U,
                FitToWidth = 1,
                FitToHeight = 0,
            };

            if (spec.AddPrinterSettings)
            {
                var settingsPart = worksheetPart.AddNewPart<SpreadsheetPrinterSettingsPart>();
                using (var settingsStream = new MemoryStream("架空のプリンター設定"u8.ToArray()))
                {
                    settingsPart.FeedData(settingsStream);
                }

                pageSetup.Id = worksheetPart.GetIdOfPart(settingsPart);
            }

            children.Add(pageSetup);
        }

        if (spec.AddHeaderFooter)
        {
            var headerFooter = new HeaderFooter(
                new OddHeader(spec.AddHeaderFooterImageCode ? "&L&G架空ヘッダー" : "&L架空ヘッダー"),
                new OddFooter("&C架空フッター"));

            if (spec.AddDistinctHeaderFooter)
            {
                headerFooter.DifferentOddEven = true;
                headerFooter.DifferentFirst = true;
                headerFooter.ScaleWithDoc = false;
                headerFooter.AlignWithMargins = false;
                headerFooter.Append(new EvenHeader("&R架空偶数ヘッダー"));
                headerFooter.Append(new EvenFooter("&L架空偶数フッター"));
                headerFooter.Append(new FirstHeader("&C架空先頭ヘッダー"));
                headerFooter.Append(new FirstFooter("&R架空先頭フッター"));
            }

            children.Add(headerFooter);
        }

        if (spec.AddRowBreaks)
        {
            children.Add(new RowBreaks(
                new Break { Id = spec.AddBrokenBreak ? 0U : 2U, Max = 16383U, ManualPageBreak = true },
                new Break { Id = 5U, Max = 16383U, ManualPageBreak = true })
            {
                Count = 2U,
                ManualBreakCount = 2U,
            });
        }

        if (spec.AddColumnBreaks)
        {
            children.Add(new ColumnBreaks(new Break { Id = 2U, Max = 1048575U, ManualPageBreak = true })
            {
                Count = 1U,
                ManualBreakCount = 1U,
            });
        }


        if (spec.AddX14DataValidation)
        {
            // Office 2010 以降の拡張入力規則。extLst の中に x14 名前空間で入る。
            children.Add(new WorksheetExtensionList(
                new WorksheetExtension(
                    new DocumentFormat.OpenXml.Office2010.Excel.DataValidations(
                        new DocumentFormat.OpenXml.Office2010.Excel.DataValidation(
                            new DocumentFormat.OpenXml.Office2010.Excel.DataValidationForumla1(
                                new DocumentFormat.OpenXml.Office.Excel.Formula("Sheet2!$A$1:$A$3")),
                            new DocumentFormat.OpenXml.Office.Excel.ReferenceSequence("A1:A5"))
                        {
                            Type = DataValidationValues.List,
                        })
                    { Count = 1U })
                {
                    Uri = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}",
                }));
        }

        var worksheet = new Worksheet();
        foreach (var child in children)
        {
            worksheet.Append(child);
        }

        worksheetPart.Worksheet = worksheet;

        if (spec.AddHyperlink || spec.Hyperlinks.Length > 0)
        {
            var hyperlinks = new Hyperlinks();

            if (spec.AddHyperlink)
            {
                var relationship = worksheetPart.AddHyperlinkRelationship(
                    new Uri("https://example.invalid/", UriKind.Absolute), isExternal: true);
                hyperlinks.Append(new Hyperlink { Reference = "A1", Id = relationship.Id });
            }

            foreach (var link in spec.Hyperlinks)
            {
                var element = new Hyperlink { Reference = link.Reference };

                if (link.UseDanglingRelationshipId)
                {
                    element.Id = "rIdMissing";
                }
                else if (link.UseInternalRelationship)
                {
                    // 外部ではない relationship を指す(本来ハイパーリンクでは使わない形)。
                    var relationship = worksheetPart.AddHyperlinkRelationship(
                        new Uri("sheet-other.xml", UriKind.Relative), isExternal: false);
                    element.Id = relationship.Id;
                }
                else if (link.ExternalTarget is { } target)
                {
                    var relationship = worksheetPart.AddHyperlinkRelationship(
                        new Uri(target, link.ExternalTargetIsRelative ? UriKind.Relative : UriKind.Absolute),
                        isExternal: true);
                    element.Id = relationship.Id;
                }

                if (link.Location is { } location)
                {
                    element.Location = location;
                }

                if (link.Tooltip is { } tooltip)
                {
                    element.Tooltip = tooltip;
                }

                if (link.Display is { } display)
                {
                    element.Display = display;
                }

                hyperlinks.Append(element);
            }

            worksheet.Append(hyperlinks);
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

        if (spec.AddHeaderFooterDrawing)
        {
            var headerFooterDrawingPart = worksheetPart.AddNewPart<DrawingsPart>();
            headerFooterDrawingPart.WorksheetDrawing =
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
            worksheet.Append(new DrawingHeaderFooter { Id = worksheetPart.GetIdOfPart(headerFooterDrawingPart) });
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

        if (spec.IsVeryHidden)
        {
            sheet.State = SheetStateValues.VeryHidden;
        }
        else if (spec.IsHidden)
        {
            sheet.State = SheetStateValues.Hidden;
        }

        workbookPart.Workbook!.Sheets!.Append(sheet);
    }

    private static DataValidation BuildDataValidation(TestDataValidation spec, bool addUnknownAttribute)
    {
        var validation = new DataValidation
        {
            SequenceOfReferences = new ListValue<StringValue> { InnerText = spec.Sqref },
        };

        if (spec.Type is { Length: > 0 } type)
        {
            validation.Type = new EnumValue<DataValidationValues> { InnerText = type };
        }

        if (spec.Operator is { } op)
        {
            validation.Operator = new EnumValue<DataValidationOperatorValues> { InnerText = op };
        }

        if (spec.Formula1 is { } formula1)
        {
            validation.Formula1 = new Formula1(formula1);
        }

        if (spec.Formula2 is { } formula2)
        {
            validation.Formula2 = new Formula2(formula2);
        }

        if (spec.AllowBlank)
        {
            validation.AllowBlank = true;
        }

        if (spec.ShowDropDown)
        {
            validation.ShowDropDown = true;
        }

        if (spec.ShowInputMessage)
        {
            validation.ShowInputMessage = true;
        }

        if (spec.ShowErrorMessage)
        {
            validation.ShowErrorMessage = true;
        }

        if (spec.ErrorStyle is { } errorStyle)
        {
            validation.ErrorStyle = new EnumValue<DataValidationErrorStyleValues> { InnerText = errorStyle };
        }

        if (spec.ImeMode is { } imeMode)
        {
            validation.ImeMode = new EnumValue<DataValidationImeModeValues> { InnerText = imeMode };
        }

        if (spec.PromptTitle is { } promptTitle)
        {
            validation.PromptTitle = promptTitle;
        }

        if (spec.Prompt is { } prompt)
        {
            validation.Prompt = prompt;
        }

        if (spec.ErrorTitle is { } errorTitle)
        {
            validation.ErrorTitle = errorTitle;
        }

        if (spec.Error is { } error)
        {
            validation.Error = error;
        }

        if (addUnknownAttribute)
        {
            validation.SetAttribute(new OpenXmlAttribute("ebt", "unknown", "urn:fictional:test", "1"));
        }

        return validation;
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
