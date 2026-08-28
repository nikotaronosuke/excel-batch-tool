using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBatchTool.Core.Aggregation;

/// <summary>
/// 複数の Source Workbook の書式を 1 つの出力 Stylesheet へまとめる。
/// Workbook ごとに Stylesheet は独立しているので、StyleIndex をそのまま持ち込まず、
/// Source 単位に「元の StyleIndex → 出力の StyleIndex」の対応表を作って変換する。
/// 重複除去(dedup)はしない。安全性と追いやすさを優先する。
/// </summary>
internal sealed class OutputStylesheetBuilder
{
    private const uint FirstCustomNumberFormatId = 164;

    private readonly List<Font> _fonts = [];
    private readonly List<Fill> _fills = [];
    private readonly List<Border> _borders = [];
    private readonly List<CellFormat> _cellFormats = [];
    private readonly List<NumberingFormat> _numberingFormats = [];
    private readonly Dictionary<string, uint[]> _mapsBySource = new(StringComparer.OrdinalIgnoreCase);

    private uint _nextCustomNumberFormatId = FirstCustomNumberFormatId;

    public OutputStylesheetBuilder()
    {
        // 出力側の既定。Excel は fills[0]=none / fills[1]=gray125 を前提にする。
        _fonts.Add(new Font(new FontSize { Val = 11D }));
        _fills.Add(new Fill(new PatternFill { PatternType = PatternValues.None }));
        _fills.Add(new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
        _borders.Add(new Border());
        _cellFormats.Add(new CellFormat
        {
            NumberFormatId = 0U,
            FontId = 0U,
            FillId = 0U,
            BorderId = 0U,
            FormatId = 0U,
        });
    }

    /// <summary>
    /// Source Workbook の書式を出力へ取り込み、「元の StyleIndex → 出力の StyleIndex」表を返す。
    /// 同じ Source を 2 回渡しても取り込みは 1 回だけ行う。
    /// </summary>
    public uint[] AddSource(string sourceKey, WorkbookPart sourceWorkbookPart)
    {
        if (_mapsBySource.TryGetValue(sourceKey, out var cached))
        {
            return cached;
        }

        var stylesheet = sourceWorkbookPart.WorkbookStylesPart?.Stylesheet;
        if (stylesheet is null)
        {
            var empty = Array.Empty<uint>();
            _mapsBySource[sourceKey] = empty;
            return empty;
        }

        var numberFormatMap = AddNumberingFormats(stylesheet);

        var sourceFonts = stylesheet.Fonts?.Elements<Font>().ToList() ?? [];
        var fontOffset = (uint)_fonts.Count;
        _fonts.AddRange(sourceFonts.Select(font => (Font)font.CloneNode(true)));

        var sourceFills = stylesheet.Fills?.Elements<Fill>().ToList() ?? [];
        var fillOffset = (uint)_fills.Count;
        _fills.AddRange(sourceFills.Select(fill => (Fill)fill.CloneNode(true)));

        var sourceBorders = stylesheet.Borders?.Elements<Border>().ToList() ?? [];
        var borderOffset = (uint)_borders.Count;
        _borders.AddRange(sourceBorders.Select(border => (Border)border.CloneNode(true)));

        var sourceFormats = stylesheet.CellFormats?.Elements<CellFormat>().ToList() ?? [];
        var map = new uint[sourceFormats.Count];

        for (var index = 0; index < sourceFormats.Count; index++)
        {
            var format = (CellFormat)sourceFormats[index].CloneNode(true);

            format.FontId = Shift(format.FontId?.Value, sourceFonts.Count, fontOffset, fallback: 0U);
            format.FillId = Shift(format.FillId?.Value, sourceFills.Count, fillOffset, fallback: 0U);
            format.BorderId = Shift(format.BorderId?.Value, sourceBorders.Count, borderOffset, fallback: 0U);

            var numberFormatId = format.NumberFormatId?.Value ?? 0U;
            format.NumberFormatId = numberFormatMap.TryGetValue(numberFormatId, out var mapped)
                ? mapped
                : numberFormatId;

            // 名前付きスタイルは取り込まないので、すべて既定の cellStyleXf を指す。
            format.FormatId = 0U;

            map[index] = (uint)_cellFormats.Count;
            _cellFormats.Add(format);
        }

        _mapsBySource[sourceKey] = map;
        return map;
    }

    /// <summary>元の StyleIndex を出力の StyleIndex へ変換する。範囲外は既定(0)。</summary>
    public static uint MapStyleIndex(uint[] map, uint? sourceStyleIndex)
        => sourceStyleIndex is { } index && index < map.Length ? map[index] : 0U;

    public Stylesheet Build()
    {
        var stylesheet = new Stylesheet();

        if (_numberingFormats.Count > 0)
        {
            stylesheet.NumberingFormats = new NumberingFormats(_numberingFormats.Select(f => f.CloneNode(true)))
            {
                Count = (uint)_numberingFormats.Count,
            };
        }

        stylesheet.Fonts = new Fonts(_fonts.Select(f => f.CloneNode(true))) { Count = (uint)_fonts.Count };
        stylesheet.Fills = new Fills(_fills.Select(f => f.CloneNode(true))) { Count = (uint)_fills.Count };
        stylesheet.Borders = new Borders(_borders.Select(b => b.CloneNode(true))) { Count = (uint)_borders.Count };

        stylesheet.CellStyleFormats = new CellStyleFormats(
            new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U })
        {
            Count = 1U,
        };

        stylesheet.CellFormats = new CellFormats(_cellFormats.Select(f => f.CloneNode(true)))
        {
            Count = (uint)_cellFormats.Count,
        };

        stylesheet.CellStyles = new CellStyles(
            new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U })
        {
            Count = 1U,
        };

        return stylesheet;
    }

    private Dictionary<uint, uint> AddNumberingFormats(Stylesheet stylesheet)
    {
        var map = new Dictionary<uint, uint>();
        foreach (var format in stylesheet.NumberingFormats?.Elements<NumberingFormat>() ?? [])
        {
            if (format.NumberFormatId?.Value is not { } sourceId || format.FormatCode?.Value is not { } code)
            {
                continue;
            }

            // 組み込み ID(164 未満)はそのまま使える。ユーザー定義だけ採番し直す。
            if (sourceId < FirstCustomNumberFormatId)
            {
                continue;
            }

            var newId = _nextCustomNumberFormatId++;
            map[sourceId] = newId;
            _numberingFormats.Add(new NumberingFormat { NumberFormatId = newId, FormatCode = code });
        }

        return map;
    }

    private static uint Shift(uint? sourceId, int sourceCount, uint offset, uint fallback)
        => sourceId is { } id && id < (uint)sourceCount ? offset + id : fallback;
}
