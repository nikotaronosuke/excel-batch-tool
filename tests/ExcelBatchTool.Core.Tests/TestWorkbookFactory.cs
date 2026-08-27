using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Charts = DocumentFormat.OpenXml.Drawing.Charts;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// テスト用の .xlsx を架空データのみで生成する。
/// 実業務データは一切使用しない。
/// (検出テスト用の最小構成であり、Excel での見た目までは保証しない。)
/// </summary>
internal static class TestWorkbookFactory
{
    /// <summary>1x1 透明 PNG(生成データ)。</summary>
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    /// <summary>値のみの通常ブックを作成する。</summary>
    public static void CreateNormal(string path)
        => Create(path, workbookPart => AddValueSheet(workbookPart, "データ", 1, Rows3x2()));

    /// <summary>複数シートのブックを作成する。</summary>
    public static void CreateMultiSheet(string path, params string[] sheetNames)
        => Create(path, workbookPart =>
        {
            uint sheetId = 1;
            foreach (var name in sheetNames)
            {
                AddValueSheet(workbookPart, name, sheetId++, Rows3x2());
            }
        });

    /// <summary>数式入りブックを作成する。</summary>
    public static void CreateWithFormulas(string path)
        => Create(path, workbookPart =>
        {
            var worksheetPart = AddValueSheet(workbookPart, "計算", 1, Rows3x2());
            var sheetData = worksheetPart.Worksheet!.GetFirstChild<SheetData>()!;
            var row = sheetData.Elements<Row>().First();
            row.Append(new Cell
            {
                CellReference = "C1",
                CellFormula = new CellFormula("A1&B1"),
                DataType = CellValues.String,
                CellValue = new CellValue("架空1架空2"),
            });
        });

    /// <summary>結合セル入りブックを作成する。</summary>
    public static void CreateWithMergedCells(string path)
        => Create(path, workbookPart =>
        {
            var worksheetPart = AddValueSheet(workbookPart, "結合", 1, Rows3x2());
            var sheetData = worksheetPart.Worksheet!.GetFirstChild<SheetData>()!;
            worksheetPart.Worksheet!.InsertAfter(
                new MergeCells(new MergeCell { Reference = "A1:B2" }), sheetData);
        });

    /// <summary>シート保護入りブックを作成する。</summary>
    public static void CreateWithSheetProtection(string path)
        => Create(path, workbookPart =>
        {
            var worksheetPart = AddValueSheet(workbookPart, "保護", 1, Rows3x2());
            var sheetData = worksheetPart.Worksheet!.GetFirstChild<SheetData>()!;
            worksheetPart.Worksheet!.InsertAfter(
                new SheetProtection { Sheet = true, Objects = true, Scenarios = true }, sheetData);
        });

    /// <summary>グラフ(最小構成の ChartPart)入りブックを作成する。</summary>
    public static void CreateWithChart(string path)
        => Create(path, workbookPart =>
        {
            var worksheetPart = AddValueSheet(workbookPart, "グラフ元", 1, Rows3x2());
            var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();

            var chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new Charts.ChartSpace(new Charts.Chart(new Charts.PlotArea()));

            worksheetPart.Worksheet!.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
        });

    /// <summary>画像(ImagePart)入りブックを作成する。</summary>
    public static void CreateWithImage(string path)
        => Create(path, workbookPart =>
        {
            var worksheetPart = AddValueSheet(workbookPart, "画像", 1, Rows3x2());
            var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();

            var imagePart = drawingsPart.AddImagePart("image/png");
            using (var stream = new MemoryStream(TinyPng))
            {
                imagePart.FeedData(stream);
            }

            worksheetPart.Worksheet!.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
        });

    /// <summary>外部参照(External Link)入りブックを作成する。</summary>
    public static void CreateWithExternalLink(string path)
        => Create(path, workbookPart =>
        {
            AddValueSheet(workbookPart, "参照元", 1, Rows3x2());

            var externalPart = workbookPart.AddNewPart<ExternalWorkbookPart>();
            var relationship = externalPart.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
                new Uri("fictional-external.xlsx", UriKind.Relative));
            externalPart.ExternalLink = new ExternalLink(new ExternalBook { Id = relationship.Id });

            workbookPart.Workbook!.AppendChild(new ExternalReferences(
                new ExternalReference { Id = workbookPart.GetIdOfPart(externalPart) }));
        });

    /// <summary>zip として不正な(壊れた)ファイルを .xlsx 拡張子で作成する。</summary>
    public static void CreateCorrupt(string path)
        => File.WriteAllBytes(path, "これは有効な xlsx ではない架空のバイト列です。"u8.ToArray());

    private static void Create(string path, Action<WorkbookPart> build)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        workbookPart.Workbook!.AppendChild(new Sheets());
        build(workbookPart);
    }

    /// <summary>値セルのみのワークシートを追加する。</summary>
    private static WorksheetPart AddValueSheet(
        WorkbookPart workbookPart, string sheetName, uint sheetId, string[][] rows)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        var sheetData = new SheetData();
        var maxColumnCount = 0;
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var row = new Row { RowIndex = (uint)(rowIndex + 1) };
            var cells = rows[rowIndex];
            maxColumnCount = Math.Max(maxColumnCount, cells.Length);
            for (var columnIndex = 0; columnIndex < cells.Length; columnIndex++)
            {
                row.Append(new Cell
                {
                    CellReference = $"{CellRangeParser.ColumnIndexToLetters(columnIndex + 1)}{rowIndex + 1}",
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Text(cells[columnIndex])),
                });
            }

            sheetData.Append(row);
        }

        var dimension = new SheetDimension
        {
            Reference = rows.Length == 0
                ? "A1"
                : $"A1:{CellRangeParser.ColumnIndexToLetters(Math.Max(1, maxColumnCount))}{rows.Length}",
        };

        worksheetPart.Worksheet = new Worksheet(dimension, sheetData);

        workbookPart.Workbook!.Sheets!.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = sheetName,
        });

        return worksheetPart;
    }

    private static string[][] Rows3x2() =>
    [
        ["架空1", "架空2"],
        ["架空3", "架空4"],
        ["架空5", "架空6"],
    ];
}
