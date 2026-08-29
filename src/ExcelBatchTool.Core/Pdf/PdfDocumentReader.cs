using Tabula;
using Tabula.Extractors;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;

namespace ExcelBatchTool.Core.Pdf;

/// <summary>1 ページを見たときの種類。</summary>
internal enum PdfPageKind
{
    Text,
    Table,

    /// <summary>文字情報が実用量に満たない(画像だけのページ)。</summary>
    ImageOnly,
}

/// <summary>読み取れなかった理由。画面には例外の内容をそのまま出さない。</summary>
internal sealed record PdfOpenFailure(string Message);

/// <summary>
/// PDF を読む。データ元は読み取りのみで、書き換えは一切しない。
///
/// 文字は PdfPig、表の構造は Tabula の格子を使うが、
/// **セルの文字は Tabula の戻り値を信用せず PdfPig の letter から詰め直す**
/// (Phase 2F-R で、連続する同一文字が潰れる問題を実測したため)。
/// 罫線のない表は、ヘッダー行の文字位置から列を決める自前の再構成を使う。
/// </summary>
internal static class PdfDocumentReader
{
    /// <summary>ページの種類を決めるための計測値。</summary>
    internal sealed record PageProfile(
        int Page,
        int Letters,
        double ImageCoverage,
        int HorizontalRulings,
        int VerticalRulings,
        bool ColumnsAligned)
    {
        /// <summary>
        /// 文字がまったく無いか、「文字がごく少ない上にページの大半が画像」なら
        /// 画像だけのページとみなす。文字が少ないだけのページ(表紙・区切りなど)は
        /// 画像扱いにしない。
        /// </summary>
        public bool IsImageOnly => Letters == 0
            || (Letters < PdfReadDefaults.MinLettersPerTextPage && ImageCoverage > 0.5);

        /// <summary>
        /// 縦横の罫線が組み合わさって格子になっているか。
        /// 記入欄の下線のように横線だけが並ぶページは、格子とみなさない。
        /// </summary>
        public bool HasGrid => HorizontalRulings >= 3 && VerticalRulings >= 3;

        /// <summary>
        /// 表とみなすのは「罫線が格子になっている」か「列の位置がそろった行が続く」ページ。
        /// 罫線のない表(位置だけで列が分かれている表)も拾えるようにしている。
        /// </summary>
        public PdfPageKind Kind => IsImageOnly
            ? PdfPageKind.ImageOnly
            : HasGrid || ColumnsAligned ? PdfPageKind.Table : PdfPageKind.Text;
    }

    /// <summary>PDF を開いて各ページを見た結果。</summary>
    internal sealed record PdfScan(
        IReadOnlyList<PageProfile> Pages,
        PdfDocumentKind Kind,
        bool AnyTablePage);

    /// <summary>PDF を開き、種類を判定する。開けないときは理由を返す。</summary>
    public static (PdfScan? Scan, PdfOpenFailure? Failure) Inspect(string filePath)
        => Open(filePath, document =>
        {
            var pages = new List<PageProfile>();
            for (var number = 1; number <= document.NumberOfPages; number++)
            {
                var page = document.GetPage(number);
                var area = page.Width * page.Height;
                var coverage = area <= 0
                    ? 0
                    : page.GetImages().Sum(image => image.BoundingBox.Width * image.BoundingBox.Height) / area;

                var boxes = page.Paths
                    .Select(path => path.GetBoundingRectangle())
                    .Where(box => box is not null)
                    .Select(box => box!.Value)
                    .ToList();

                var horizontal = boxes.Count(box => box.Height < 2 && box.Width > 100);
                var vertical = boxes.Count(box => box.Width < 2 && box.Height > 20);

                pages.Add(new PageProfile(
                    number, page.Letters.Count, coverage, horizontal, vertical,
                    HasAlignedColumns(page)));
            }

            var textPages = pages.Count(page => page.Kind != PdfPageKind.ImageOnly);
            var kind = pages.Count == 0 ? PdfDocumentKind.Unknown
                : textPages == 0 ? PdfDocumentKind.Scan
                : textPages < pages.Count ? PdfDocumentKind.Mixed
                : pages.Any(page => page.Kind == PdfPageKind.Table)
                    ? PdfDocumentKind.Table
                    : PdfDocumentKind.Text;

            return new PdfScan(pages, kind, pages.Any(page => page.Kind == PdfPageKind.Table));
        });

    /// <summary>
    /// 罫線がなくても、列の位置がそろった行が続いていれば表とみなす。
    ///
    /// 各行を「6pt 以上の空きで切った塊」に分け、同じ塊数(2 以上)を持つ行が
    /// 全体の 6 割以上あり、かつ塊の左端がほぼ同じ位置に並んでいるかを見る。
    /// </summary>
    private static bool HasAlignedColumns(Page page)
    {
        var lines = LineGroups(page.Letters);
        if (lines.Count < 3)
        {
            return false;
        }

        var segments = lines.Select(SegmentStarts).ToList();
        var counts = segments
            .Where(starts => starts.Count >= 2)
            .GroupBy(starts => starts.Count)
            .OrderByDescending(group => group.Count())
            .FirstOrDefault();

        if (counts is null || counts.Count() < lines.Count * 0.6)
        {
            return false;
        }

        // 同じ塊数の行どうしで、各列の左端がそろっているか。
        var aligned = counts.ToList();
        for (var column = 0; column < counts.Key; column++)
        {
            var positions = aligned.Select(starts => starts[column]).ToList();
            if (positions.Max() - positions.Min() > 6)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>1 行の中で、6pt 以上の空きで切った塊の左端。</summary>
    private static List<double> SegmentStarts(List<Letter> letters)
    {
        var starts = new List<double>();
        double? lastRight = null;

        foreach (var letter in letters)
        {
            if (lastRight is null || letter.BoundingBox.Left - lastRight > 6)
            {
                starts.Add(letter.BoundingBox.Left);
            }

            lastRight = Math.Max(lastRight ?? 0, letter.BoundingBox.Right);
        }

        return starts;
    }

    /// <summary>ベースラインでまとめた行(上から順)。</summary>
    private static List<List<Letter>> LineGroups(IReadOnlyList<Letter> letters)
        => letters
            .GroupBy(letter => Math.Round(letter.BoundingBox.Bottom / 4))
            .OrderByDescending(group => group.Key)
            .Select(group => group.OrderBy(letter => letter.BoundingBox.Left).ToList())
            .ToList();

    /// <summary>通常の文字 PDF を、ページ番号と行番号を保ったまま読む。</summary>
    public static (IReadOnlyList<PdfTextLine>? Lines, PdfOpenFailure? Failure) ReadLines(string filePath)
        => Open(filePath, document =>
        {
            var lines = new List<PdfTextLine>();
            for (var number = 1; number <= document.NumberOfPages; number++)
            {
                var page = document.GetPage(number);
                var lineNumber = 0;
                foreach (var text in GroupIntoLines(page.Letters))
                {
                    var normalized = PdfTextNormalization.Normalize(text);
                    if (normalized.Length > 0)
                    {
                        lines.Add(new PdfTextLine(number, ++lineNumber, normalized));
                    }
                }
            }

            return (IReadOnlyList<PdfTextLine>)lines;
        });

    /// <summary>表 PDF を、行 × 列の文字列へ復元する。</summary>
    public static (PdfTableResult? Table, PdfOpenFailure? Failure) ReadTable(string filePath)
    {
        var (result, failure) = Open(filePath, document =>
        {
            var pages = new List<PdfTablePage>();
            var usedRulings = false;

            for (var number = 1; number <= document.NumberOfPages; number++)
            {
                var page = document.GetPage(number);
                var rows = ReadRuledTable(document, page, number);
                if (rows is not null)
                {
                    usedRulings = true;
                }
                else
                {
                    rows = ReadHeaderGuidedTable(page);
                }

                pages.Add(new PdfTablePage(number, rows ?? []));
            }

            return new PdfTableResult(pages, usedRulings);
        });

        return (result, failure);
    }

    /// <summary>罫線のある表: 構造は Tabula、文字は PdfPig の letter から詰め直す。</summary>
    private static List<string[]>? ReadRuledTable(PdfDocument document, Page page, int number)
    {
        List<Tabula.Table> tables;
        try
        {
            var area = ObjectExtractor.Extract(document, number);
            tables = new SpreadsheetExtractionAlgorithm().Extract(area).ToList();
        }
        catch (Exception)
        {
            return null;
        }

        // 実際の格子を選ぶ。Tabula は表の外枠を「1 行 1 セル」の格子としても返すので、
        // 「短いセルが 2 つ以上並ぶ行を持つ」ものだけを候補にする
        // (外枠だけの格子は候補から外れる)。
        var table = tables
            .Where(candidate => candidate.Rows.Any(row =>
                row.Count(cell =>
                    PdfTextNormalization.Normalize(cell.GetText()).Length is > 0 and < 40) >= 2))
            .OrderByDescending(candidate => candidate.RowCount)
            .FirstOrDefault();

        if (table is null)
        {
            return null;
        }

        // Tabula は表の外枠を「表全体の高さを 1 つのセルが占める行」としても返す。
        // 行の高さを見て、通常の行の 2 倍を超えるものは外枠として落とす。
        var candidates = table.Rows
            .Select(row => (
                Height: row.Count == 0 ? 0 : row.Max(cell => cell.BoundingBox.Height),
                Cells: row
                    .Select(cell => PdfTextNormalization.Normalize(LettersInBox(page, cell.BoundingBox)))
                    .ToArray()))
            .Where(row => row.Cells.Any(cell => cell.Length > 0))
            .ToList();

        if (candidates.Count == 0)
        {
            return null;
        }

        var normalHeight = candidates.Min(row => row.Height);
        var rows = candidates
            .Where(row => normalHeight <= 0 || row.Height <= normalHeight * 2)
            .Select(row => row.Cells)
            .ToList();

        return rows.Count > 0 ? rows : null;
    }

    /// <summary>罫線のない表: ヘッダー行の文字位置から列の左端を決め、以降の行を位置で割り当てる。</summary>
    private static List<string[]>? ReadHeaderGuidedTable(Page page)
    {
        var rows = LineGroups(page.Letters);

        if (rows.Count < 2)
        {
            return null;
        }

        // 1 行目 = ヘッダー。文字の途切れ(> 6pt)を列の境目とみなす。
        var columns = SegmentStarts(rows[0]);

        if (columns.Count < 2)
        {
            return null;
        }

        return rows
            .Select(letters =>
            {
                var cells = new string[columns.Count];
                foreach (var group in letters.GroupBy(letter => ColumnOf(columns, letter)))
                {
                    cells[group.Key] = PdfTextNormalization.Normalize(string.Concat(
                        group.OrderBy(letter => letter.BoundingBox.Left).Select(letter => letter.Value)));
                }

                return cells.Select(cell => cell ?? string.Empty).ToArray();
            })
            .Where(row => row.Any(cell => cell.Length > 0))
            .ToList();
    }

    private static int ColumnOf(List<double> columns, Letter letter)
    {
        var x = letter.BoundingBox.Left + 0.5;
        for (var index = columns.Count - 1; index >= 0; index--)
        {
            if (x >= columns[index] - 1)
            {
                return index;
            }
        }

        return 0;
    }

    /// <summary>セル矩形の中にある letter を、左から順に並べて文字列にする。</summary>
    private static string LettersInBox(Page page, PdfRectangle box)
    {
        var inside = page.Letters
            .Where(letter =>
            {
                var x = (letter.BoundingBox.Left + letter.BoundingBox.Right) / 2;
                var y = (letter.BoundingBox.Top + letter.BoundingBox.Bottom) / 2;
                return x >= box.Left - 2 && x <= box.Right + 2
                    && y >= box.Bottom - 2 && y <= box.Top + 2;
            })
            .OrderBy(letter => letter.BoundingBox.Left);

        return string.Concat(inside.Select(letter => letter.Value));
    }

    /// <summary>ベースラインでまとめて 1 行にする(段組は左から右の順に並ぶ)。</summary>
    private static IEnumerable<string> GroupIntoLines(IReadOnlyList<Letter> letters)
        => LineGroups(letters).Select(line => string.Concat(line.Select(letter => letter.Value)));

    /// <summary>
    /// PDF を開いて読む。パスワード付き・壊れている・ページが無い等は、
    /// 例外の内容ではなく利用者向けの文言にして返す。
    /// </summary>
    private static (T? Value, PdfOpenFailure? Failure) Open<T>(string filePath, Func<PdfDocument, T> read)
        where T : class
    {
        if (!File.Exists(filePath))
        {
            return (null, new PdfOpenFailure("PDF ファイルが見つかりません。"));
        }

        try
        {
            // ClipPaths を有効にしないと、Tabula が罫線を格子として読み取れない。
            using var document = PdfDocument.Open(filePath, new ParsingOptions { ClipPaths = true });

            if (document.NumberOfPages == 0)
            {
                return (null, new PdfOpenFailure("この PDF にはページがありません。"));
            }

            if (document.NumberOfPages > PdfReadDefaults.MaxPages)
            {
                return (null, new PdfOpenFailure(
                    $"ページ数({document.NumberOfPages:N0})が動作を確認した範囲"
                        + $"({PdfReadDefaults.MaxPages:N0} ページ)を超えています。"));
            }

            return (read(document), null);
        }
        catch (Exception ex) when (IsEncrypted(ex))
        {
            return (null, new PdfOpenFailure(
                "この PDF はパスワードで保護されているため読み取れません。"
                    + "保護を外したファイルを選んでください。"));
        }
        catch (Exception ex) when (ex is PdfDocumentFormatException or InvalidOperationException
            or IndexOutOfRangeException or ArgumentException or NotSupportedException or IOException
            or NullReferenceException or FormatException or OverflowException)
        {
            return (null, new PdfOpenFailure(
                "この PDF を読み取れません。ファイルが壊れているか、"
                    + "現在のバージョンでは扱えない形式の可能性があります。"));
        }
    }

    /// <summary>
    /// パスワード保護による失敗かどうか。PdfPig の暗号化例外は版によって型が変わるため、
    /// 型ではなく名前と文言で判定する(利用者には保護の話だけを伝えたい)。
    /// </summary>
    private static bool IsEncrypted(Exception exception)
    {
        for (var current = exception; current is not null; current = current.InnerException)
        {
            if (current.GetType().Name.Contains("Encrypt", StringComparison.OrdinalIgnoreCase)
                || current.Message.Contains("password", StringComparison.OrdinalIgnoreCase)
                || current.Message.Contains("encrypt", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }
}

/// <summary>表 PDF の 1 ページ分。</summary>
internal sealed record PdfTablePage(int Page, IReadOnlyList<string[]> Rows);

/// <summary>表 PDF の読み取り結果。</summary>
internal sealed record PdfTableResult(IReadOnlyList<PdfTablePage> Pages, bool FromRulings);
