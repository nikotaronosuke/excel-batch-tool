using ExcelBatchTool.Ocr;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Sdcb.PaddleOCR.Models;
using Sdcb.PaddleOCR.Models.Online;

// Offline OCR Pack を組み立てる。
//   OcrPackBuilder <published-ocr-dir> <output-pack-dir>
//
// モデルの取得はここでだけ行う。出来上がった Pack は完全に自己完結し、
// 製品の実行時に通信も追加インストールも要らない。
if (args.Length < 2)
{
    Console.WriteLine("usage: OcrPackBuilder <published-ocr-dir> <output-pack-dir>");
    return 1;
}

var source = Path.GetFullPath(args[0]);
var target = Path.GetFullPath(args[1]);

if (Directory.Exists(target))
{
    Directory.Delete(target, recursive: true);
}

Directory.CreateDirectory(target);

// 1. 公開済みの OCR 実装をコピーする。
//    ExcelBatchTool.Core は本体側の DLL を使うので Pack へは入れない。
var skipped = new[]
{
    "ExcelBatchTool.Core.dll",
    "ExcelBatchTool.Core.pdb",

    // 動画入出力は使わない。OpenCvSharp が同梱してくるだけなので外す(27MB)。
    "opencv_videoio_ffmpeg4130_64.dll",
};
var copied = 0;
foreach (var file in Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories))
{
    var relative = Path.GetRelativePath(source, file);
    if (skipped.Contains(Path.GetFileName(file), StringComparer.OrdinalIgnoreCase)
        || relative.EndsWith(".pdb", StringComparison.OrdinalIgnoreCase))
    {
        continue;
    }

    var destination = Path.Combine(target, relative);
    Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
    File.Copy(file, destination, overwrite: true);
    copied++;
}

Console.WriteLine($"copied {copied} files from the published OCR implementation");

// 2. モデルを取り出す。
//
//    どのモデルを入れるかは引数で選べる(既定 v5japan)。Phase 2F-B3 で、
//    「今までの v4 二重読み」と「PP-OCRv5 + 日本語 v4」を **同じ製品経路のまま**
//    比べるために、Pack を組み替えるだけで構成を変えられるようにした。
//
//    Paddle Inference 3.3.1 では PP-OCRv5 と japan_PP-OCRv4 が同じプロセスへ
//    両方読み込める(2F-B1 の「同居できない」は 3.0.x の話で、3.3.1 では再現しない)。
//    そのため別プロセスも IPC も要らない。
var models = Path.Combine(target, "models");
var config = args.Length > 2 ? args[2] : "v5japan";

CopyModel(OnlineDetectionModel.ChineseV4.DownloadAsync().GetAwaiter().GetResult().DirectoryPath,
    Path.Combine(models, "det"));
CopyModel(OnlineClassificationModel.ChineseMobileV2.DownloadAsync().GetAwaiter().GetResult().DirectoryPath,
    Path.Combine(models, "cls"));

OcrModelSet set;
switch (config)
{
    case "v4dual":
        WriteRecognitionModel(LocalDictOnlineRecognizationModel.ChineseV4,
            Path.Combine(models, "rec-multi"));
        WriteRecognitionModel(LocalDictOnlineRecognizationModel.JapanV4,
            Path.Combine(models, "rec-japan"));
        set = new OcrModelSet
        {
            PrimaryName = "ch_PP-OCRv4_rec", PrimaryVersion = 4,
            SecondaryName = "japan_PP-OCRv4_rec", SecondaryVersion = 4,
            Runtime = "Paddle Inference 3.3.1",
        };
        break;

    case "v5only":
        WriteFullRecognitionModel(OnlineFullModels.ChineseV5, Path.Combine(models, "rec-multi"));
        WriteFullRecognitionModel(OnlineFullModels.ChineseV5, Path.Combine(models, "rec-japan"));
        set = new OcrModelSet
        {
            PrimaryName = "PP-OCRv5_mobile_rec", PrimaryVersion = 5,
            SecondaryName = "PP-OCRv5_mobile_rec", SecondaryVersion = 5,
            Runtime = "Paddle Inference 3.3.1",
        };
        break;

    case "v5japan":
    default:
        WriteFullRecognitionModel(OnlineFullModels.ChineseV5, Path.Combine(models, "rec-multi"));
        WriteRecognitionModel(LocalDictOnlineRecognizationModel.JapanV4,
            Path.Combine(models, "rec-japan"));
        set = new OcrModelSet
        {
            PrimaryName = "PP-OCRv5_mobile_rec", PrimaryVersion = 5,
            SecondaryName = "japan_PP-OCRv4_rec", SecondaryVersion = 4,
            Runtime = "Paddle Inference 3.3.1",
        };
        break;
}

File.WriteAllText(Path.Combine(models, "models.json"), set.ToJson());
Console.WriteLine($"model set: {config} ({set.PrimaryName} + {set.SecondaryName})");

// 3. VC++ ランタイムを同梱する(利用者に別途インストールを求めない)。
//    実測で、この 3 つを必要とするのは ONNX 系の DLL だけ。
//    paddle_inference_c / openblas / OpenCvSharpExtern / pdfium / libSkiaSharp は
//    ランタイムを静的に取り込んでいるので追加は要らない。
CopyVisualCppRuntime(target);

// 4. ライセンス表示をまとめる。
File.WriteAllText(
    Path.Combine(target, "LICENSES.txt"),
    """
    Offline OCR Pack — 同梱物のライセンス

    Paddle Inference (native runtime) ........ Apache-2.0
      paddle_inference_c.dll / common.dll / phi.dll
      https://github.com/PaddlePaddle/Paddle

    PaddleOCR models .......................... Apache-2.0
      ch_PP-OCRv4_det / ch_ppocr_mobile_v2.0_cls /
      ch_PP-OCRv4_rec / japan_PP-OCRv4_rec
      https://github.com/PaddlePaddle/PaddleOCR

    ONNX Runtime .............................. MIT
      https://github.com/microsoft/onnxruntime

    Microsoft Visual C++ Runtime .............. Microsoft 再頒布可能コード
      VCRUNTIME140.dll / VCRUNTIME140_1.dll / MSVCP140.dll
      アプリと同じ場所へ置く形(app-local)で同梱しています。
      利用者に別途インストールを求めません。

    Paddle2ONNX ............................... Apache-2.0
      https://github.com/PaddlePaddle/Paddle2ONNX

    OpenBLAS .................................. BSD-3-Clause
      https://github.com/OpenMathLib/OpenBLAS

    Sdcb.PaddleOCR / Sdcb.PaddleInference ..... Apache-2.0
      https://github.com/sdcb/PaddleSharp

    OpenCvSharp (+ OpenCV) .................... Apache-2.0
      https://github.com/shimat/opencvsharp

    PDFtoImage (+ PDFium) ..................... MIT / BSD-3-Clause
      https://github.com/sungaila/PDFtoImage

    SkiaSharp (+ Skia) ........................ MIT / BSD-3-Clause
      https://github.com/mono/SkiaSharp

    YamlDotNet ................................ MIT
      Sdcb.PaddleOCR 3.x がモデルの設定を読むのに使う。
      https://github.com/aaubry/YamlDotNet

    いずれも商用配布物への同梱が認められている条件です。
    Intel MKL は使っていません(再配布条件に曖昧さが残るため)。
    推論の実行には ONNX Runtime を使い、行列演算の代替として OpenBLAS を同梱しています。
    どちらも再配布条件が明確です。

    """,
    new UTF8Encoding(true));

// 5. 目録を書く。欠損・サイズ違い・中身違いを起動前に見つけるために使う。
var files = new JsonArray();
foreach (var file in Directory.EnumerateFiles(target, "*", SearchOption.AllDirectories).Order(StringComparer.Ordinal))
{
    var relative = Path.GetRelativePath(target, file).Replace(Path.DirectorySeparatorChar, '/');
    if (relative == OcrPackNames.Manifest)
    {
        continue;
    }

    using var stream = File.OpenRead(file);
    files.Add(new JsonObject
    {
        ["path"] = relative,
        ["length"] = new FileInfo(file).Length,
        ["sha256"] = Convert.ToHexString(SHA256.HashData(stream)),
    });
}

var manifest = new JsonObject
{
    ["schemaVersion"] = 1,
    ["engineAssembly"] = "ExcelBatchTool.Ocr.dll",
    ["engineType"] = "ExcelBatchTool.Ocr.PaddleOcrEngine",
    ["description"] = "Offline OCR Pack (PP-OCRv4 multilingual + japan, Paddle 2.6.1, ONNX Runtime)",
    ["files"] = files,
};

File.WriteAllText(
    Path.Combine(target, OcrPackNames.Manifest),
    manifest.ToJsonString(new JsonSerializerOptions { WriteIndented = true }),
    new UTF8Encoding(false));

var total = Directory.EnumerateFiles(target, "*", SearchOption.AllDirectories)
    .Sum(file => new FileInfo(file).Length);

Console.WriteLine($"pack built: {files.Count} files, {total / 1024.0 / 1024.0:F1} MB");
Console.WriteLine(target);
return 0;

/// <summary>
/// VC++ ランタイムを Pack へ同梱する。
///
/// 実測では、これを必要とするのは ONNX 系の DLL(onnxruntime / paddle2onnx)だけで、
/// paddle_inference_c / openblas / OpenCvSharpExtern / pdfium / libSkiaSharp は
/// ランタイムを静的に取り込んでいる。製品本体は Windows 標準の Universal CRT だけで動く。
///
/// 取得元は、あれば Visual C++ 再頒布可能パッケージの Redist フォルダー、
/// 無ければ実行環境の System32(同じバイナリ)。
/// </summary>
static void CopyVisualCppRuntime(string target)
{
    string[] needed = ["VCRUNTIME140.dll", "VCRUNTIME140_1.dll", "MSVCP140.dll"];

    var candidates = new List<string>();
    foreach (var root in new[]
    {
        Environment.GetEnvironmentVariable("VCToolsRedistDir"),
        Environment.GetEnvironmentVariable("VCINSTALLDIR"),
    })
    {
        if (!string.IsNullOrEmpty(root) && Directory.Exists(root))
        {
            candidates.AddRange(Directory.EnumerateDirectories(root, "*CRT", SearchOption.AllDirectories));
        }
    }

    candidates.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System)));

    foreach (var name in needed)
    {
        var source = candidates
            .Select(directory => Path.Combine(directory, name))
            .FirstOrDefault(File.Exists);

        if (source is null)
        {
            throw new FileNotFoundException(
                $"{name} が見つかりません。Visual C++ 再頒布可能パッケージが要ります。");
        }

        File.Copy(source, Path.Combine(target, name), overwrite: true);
        Console.WriteLine($"vcruntime: {name} <- {Path.GetDirectoryName(source)}");
    }
}

// PP-OCRv5 は OnlineFullModels からしか取れない。辞書の書き出し方は同じ。
static void WriteFullRecognitionModel(OnlineFullModels online, string to)
{
    var full = online.DownloadAsync().GetAwaiter().GetResult();
    var rec = full.RecognizationModel;
    var root = FindModelRoot(rec);
    CopyModel(root, to);
    WriteDictionary(rec, to);
}

static string FindModelRoot(object model)
{
    // FileRecognizationModel は元のフォルダーを持っている。名前は版で変わりうるので探す。
    foreach (var name in new[] { "DirectoryPath", "RootDirectory", "ModelDir" })
    {
        var property = model.GetType().GetProperty(name);
        if (property?.GetValue(model) is string path && Directory.Exists(path))
        {
            return path;
        }
    }

    foreach (var field in model.GetType()
        .GetFields(System.Reflection.BindingFlags.Instance
            | System.Reflection.BindingFlags.NonPublic
            | System.Reflection.BindingFlags.Public))
    {
        if (field.GetValue(model) is string path && Directory.Exists(path))
        {
            return path;
        }
    }

    throw new InvalidOperationException(
        $"{model.GetType().Name} からモデルのフォルダーを取り出せませんでした。");
}

static void CopyModel(string from, string to)
{
    Directory.CreateDirectory(to);
    foreach (var file in Directory.EnumerateFiles(from))
    {
        var name = Path.GetFileName(file);

        // 推論に要らない付随ファイルは入れない(Pack を小さく保つ)。
        if (name.StartsWith("._", StringComparison.Ordinal)
            || name.EndsWith(".info", StringComparison.Ordinal))
        {
            continue;
        }

        File.Copy(file, Path.Combine(to, name), overwrite: true);
    }

    Console.WriteLine($"model: {Path.GetFileName(to)} <- {Path.GetFileName(from)}");
}

// 認識モデルは、辞書がファイルとして存在しない(パッケージが持っている)。
// Pack では実ファイルとして持たせたいので、モデルから 1 行ずつ取り出して書き出す。
static void WriteRecognitionModel(LocalDictOnlineRecognizationModel online, string to)
{
    var model = online.DownloadAsync().GetAwaiter().GetResult();
    CopyModel(online.RootDirectory, to);

    WriteDictionary(model, to);
}

// 認識モデルは辞書をファイルとして持っていない(パッケージの中にある)。
// Pack では実ファイルにしたいので、1 行ずつ取り出して書き出す。
static void WriteDictionary(object model, string to)
{
    var method = model.GetType().GetMethod("GetLabelByIndex", [typeof(int)])
        ?? throw new InvalidOperationException(
            $"{model.GetType().Name} に GetLabelByIndex がありません。");

    var labels = new List<string>();
    for (var index = 1; ; index++)
    {
        string label;
        try
        {
            label = (string)method.Invoke(model, [index])!;
        }
        catch
        {
            break;
        }

        // 末尾の 1 件は「空白」を表す固定の項目で、辞書の中身ではない。
        labels.Add(label);
    }

    if (labels.Count > 0)
    {
        labels.RemoveAt(labels.Count - 1);
    }

    File.WriteAllLines(Path.Combine(to, "dict.txt"), labels, new UTF8Encoding(false));
    Console.WriteLine($"  dict.txt: {labels.Count} entries");
}

internal static class OcrPackNames
{
    public const string Manifest = "pack.json";
}
