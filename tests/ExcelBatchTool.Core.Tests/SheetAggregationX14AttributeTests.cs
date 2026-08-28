using System.IO.Compression;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelBatchTool.Core.Aggregation;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 1B.2B2B.1: x14 入力規則の属性判定。
/// 入力で例外的に許すのは既知の xr:uid だけで、同じ名前空間でも名前が違えば Block する。
/// 未知の属性を受け入れておいて出力で黙って消す、という挙動にはしない。
/// </summary>
public sealed class SheetAggregationX14AttributeTests
{
    private const string RevisionNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
    private const string SampleUid = "{11111111-2222-3333-4444-555555555555}";

    [Fact]
    public void Aggregate_X14ValidationWithOnlyRevisionUid_IsAcceptedAndTheUidIsNotCarriedOver()
    {
        using var dir = new TempDir();
        var path = MasterWorkbook(dir, new TestX14Validation(
            "B2:B100", "商品マスタ!$A$2:$A$50", RevisionUid: SampleUid));

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "注文"), (path, "商品マスタ")).Success);

        // 規則そのものは保持される。
        Assert.Equal("'商品マスタ'!$A$2:$A$50", SingleX14ListSource(output, "注文"));

        // 出力にはリビジョン識別子を持ち込まない。
        using (var zip = ZipFile.OpenRead(output))
        {
            foreach (var entry in zip.Entries.Where(e =>
                e.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal)))
            {
                using var reader = new StreamReader(entry.Open());
                Assert.DoesNotContain(RevisionNamespace, reader.ReadToEnd(), StringComparison.Ordinal);
            }
        }

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator().Validate(document).ToList();
        Assert.True(errors.Count == 0,
            string.Join(" / ", errors.Take(5).Select(error => $"{error.Path?.XPath}: {error.Description}")));
    }

    public static TheoryData<bool, bool, bool> UnknownAttributeCases => new()
    {
        // uid なし + xr 名前空間の未知属性
        { false, true, false },
        // uid あり + xr 名前空間の未知属性
        { true, true, false },
        // 別名前空間の未知属性
        { false, false, true },
        // uid あり + 別名前空間の未知属性
        { true, false, true },
    };

    [Theory]
    [MemberData(nameof(UnknownAttributeCases))]
    public void Preview_X14ValidationWithAnUnknownAttribute_IsBlocked(
        bool withUid, bool unknownRevisionAttribute, bool unknownOtherAttribute)
    {
        using var dir = new TempDir();
        var path = MasterWorkbook(dir, new TestX14Validation(
            "B2:B100",
            "商品マスタ!$A$2:$A$50",
            RevisionUid: withUid ? SampleUid : null,
            AddUnknownRevisionAttribute: unknownRevisionAttribute,
            AddUnknownAttribute: unknownOtherAttribute));

        var preview = CreatePreview((path, "注文"), (path, "商品マスタ"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("対応していない設定"));
    }

    [Fact]
    public void Aggregate_AfterTheFix_MasterSheetListAndNamedRangeStillWork()
    {
        using var dir = new TempDir();
        var path = dir.File("注文.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "注文",
                Rows = [["注文"]],
                DataValidations = [new TestDataValidation("C2:C100", "list", Formula1: "商品一覧")],
                X14Validations =
                [
                    new TestX14Validation("B2:B100", "商品マスタ!$A$2:$A$50", RevisionUid: SampleUid),
                ],
            },
            new TestAggregationSheetSpec { Name = "商品マスタ", Rows = [["架空商品"]] },
        ],
            definedNames: [new TestDefinedName("商品一覧", "'商品マスタ'!$A$2:$A$50")]);

        var before = Snapshot(path);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "注文"), (path, "商品マスタ")).Success);

        // 入力ファイルは変わらない。
        Assert.Equal(before, Snapshot(path));

        Assert.Equal("'商品マスタ'!$A$2:$A$50", SingleX14ListSource(output, "注文"));

        var validation = Assert.Single(Worksheet(output, "注文").Descendants<DataValidation>());
        Assert.Equal("商品一覧", validation.Formula1?.Text);

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var definedName = Assert.Single(
            document.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>()
                .Where(name => name.LocalSheetId is null));
        Assert.Equal("'商品マスタ'!$A$2:$A$50", definedName.Text);
    }

    // --- helpers -------------------------------------------------------

    private static string MasterWorkbook(TempDir dir, TestX14Validation validation)
    {
        var path = dir.File("注文.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec { Name = "注文", Rows = [["注文"]], X14Validations = [validation] },
            new TestAggregationSheetSpec { Name = "商品マスタ", Rows = [["架空商品"]] },
        ]);

        return path;
    }

    private static SheetAggregationPreview CreatePreview(params (string Path, string Sheet)[] selections)
        => new SheetAggregationPlanner().CreatePreview(
            [.. selections.Select(s => new SheetSelection(s.Path, s.Sheet))]);

    private static SheetAggregationResult Aggregate(string output, params (string Path, string Sheet)[] selections)
    {
        var preview = CreatePreview(selections);
        Assert.True(preview.CanExecute,
            string.Join(" / ", preview.Blocks.Select(issue => $"{issue.Location}: {issue.Message}")));
        return new SheetAggregator().Execute(preview, output);
    }

    private static Worksheet Worksheet(string path, string sheetName)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name?.Value == sheetName);
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet!;
    }

    private static string? SingleX14ListSource(string path, string sheetName)
    {
        var validation = Assert.Single(Worksheet(path, sheetName).Descendants<X14.DataValidation>());
        return validation.DataValidationForumla1?.GetFirstChild<Xm.Formula>()?.Text;
    }

    private static (string Hash, long Length, DateTime LastWriteUtc) Snapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        var info = new FileInfo(path);
        return (hash, info.Length, info.LastWriteTimeUtc);
    }
}
