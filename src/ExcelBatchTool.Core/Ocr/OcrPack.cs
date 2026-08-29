using System.Reflection;
using System.Runtime.Loader;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelBatchTool.Core.Ocr;

/// <summary>Offline OCR Pack の中身の目録。ファイルが欠けていないか・入れ替わっていないかを見る。</summary>
public sealed record OcrPackManifest
{
    [JsonPropertyName("schemaVersion")]
    public int SchemaVersion { get; init; }

    [JsonPropertyName("engineAssembly")]
    public string EngineAssembly { get; init; } = string.Empty;

    [JsonPropertyName("engineType")]
    public string EngineType { get; init; } = string.Empty;

    [JsonPropertyName("description")]
    public string Description { get; init; } = string.Empty;

    [JsonPropertyName("files")]
    public IReadOnlyList<OcrPackFile> Files { get; init; } = [];
}

public sealed record OcrPackFile
{
    [JsonPropertyName("path")]
    public string Path { get; init; } = string.Empty;

    [JsonPropertyName("length")]
    public long Length { get; init; }

    [JsonPropertyName("sha256")]
    public string Sha256 { get; init; } = string.Empty;
}

/// <summary>OCR Pack が使えるかどうか。使えない理由は利用者向けの文言で持つ。</summary>
public sealed record OcrPackStatus(bool IsPresent, bool IsUsable, string Message, string Directory)
{
    internal OcrPackManifest? Manifest { get; init; }

    public static OcrPackStatus Missing(string directory) => new(
        false,
        false,
        "スキャン PDF を読み取るための「Offline OCR Pack」が見つかりません。"
            + "配布物の ocr フォルダーをアプリと同じ場所に置いてください。"
            + "文字情報のある PDF は OCR Pack がなくても読み取れます。",
        directory);

    public static OcrPackStatus Broken(string directory, string detail) => new(
        true,
        false,
        $"「Offline OCR Pack」を使える状態ではありません({detail})。"
            + "配布物の ocr フォルダーを入れ直してください。",
        directory);
}

/// <summary>
/// Offline OCR Pack の探索・検査・読み込み。
///
/// 製品本体は OCR ランタイムを一切参照しない。Pack が無くてもアプリは普通に起動し、
/// 文字情報のある PDF はそのまま読める。スキャン PDF を選んだときだけ、
/// ここで見つけた実装を使う。
/// </summary>
public static class OcrPack
{
    public const string DirectoryName = "ocr";

    public const string ManifestFileName = "pack.json";

    public const int SupportedSchemaVersion = 1;

    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNameCaseInsensitive = false,
        ReadCommentHandling = JsonCommentHandling.Disallow,
    };

    /// <summary>差し替えできるようにしておく(テストで Pack の場所を指定する)。</summary>
    public static string DefaultDirectory
        => Path.Combine(AppContext.BaseDirectory, DirectoryName);

    /// <summary>
    /// Pack の状態を調べる。ここで壊れているものを見つけて文言にし、
    /// native の読み込みで不可解な落ち方をしないようにする。
    /// </summary>
    public static OcrPackStatus Inspect(string? directory = null)
    {
        var root = directory ?? DefaultDirectory;

        if (!Directory.Exists(root))
        {
            return OcrPackStatus.Missing(root);
        }

        var manifestPath = Path.Combine(root, ManifestFileName);
        if (!File.Exists(manifestPath))
        {
            return OcrPackStatus.Broken(root, $"{ManifestFileName} がありません");
        }

        OcrPackManifest? manifest;
        try
        {
            manifest = JsonSerializer.Deserialize<OcrPackManifest>(
                File.ReadAllText(manifestPath), Options);
        }
        catch (JsonException)
        {
            return OcrPackStatus.Broken(root, $"{ManifestFileName} の内容を読み取れません");
        }

        if (manifest is null)
        {
            return OcrPackStatus.Broken(root, $"{ManifestFileName} の内容が空です");
        }

        if (manifest.SchemaVersion != SupportedSchemaVersion)
        {
            return OcrPackStatus.Broken(
                root,
                $"この版のアプリが扱えない形式です(必要 {SupportedSchemaVersion} / 実際 {manifest.SchemaVersion})");
        }

        if (manifest.EngineAssembly.Length == 0 || manifest.EngineType.Length == 0)
        {
            return OcrPackStatus.Broken(root, $"{ManifestFileName} に必要な項目がありません");
        }

        if (manifest.Files.Count == 0)
        {
            return OcrPackStatus.Broken(root, $"{ManifestFileName} にファイルの一覧がありません");
        }

        if (Verify(root, manifest) is { } problem)
        {
            return OcrPackStatus.Broken(root, problem);
        }

        return new OcrPackStatus(true, true, "OCR Pack を使えます。", root) { Manifest = manifest };
    }

    /// <summary>目録どおりのファイルが揃っているか(欠損・サイズ違い・中身違い)。</summary>
    private static string? Verify(string root, OcrPackManifest manifest)
    {
        foreach (var entry in manifest.Files)
        {
            if (entry.Path.Length == 0
                || Path.IsPathRooted(entry.Path)
                || entry.Path.Contains("..", StringComparison.Ordinal))
            {
                return "ファイルの一覧に使えない指定があります";
            }

            var path = Path.Combine(root, entry.Path.Replace('/', Path.DirectorySeparatorChar));
            if (!File.Exists(path))
            {
                return $"{entry.Path} がありません";
            }

            var info = new FileInfo(path);
            if (info.Length != entry.Length)
            {
                return $"{entry.Path} の大きさが違います";
            }

            if (entry.Sha256.Length > 0 && !HashMatches(path, entry.Sha256))
            {
                return $"{entry.Path} の内容が違います";
            }
        }

        return null;
    }

    private static bool HashMatches(string path, string expected)
    {
        try
        {
            using var stream = new FileStream(
                path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 20);
            var hash = Convert.ToHexString(SHA256.HashData(stream));
            return string.Equals(hash, expected, StringComparison.OrdinalIgnoreCase);
        }
        catch (IOException)
        {
            return false;
        }
    }

    /// <summary>検査を通った Pack から OCR の実体を読み込む。</summary>
    public static IOcrEngine Load(OcrPackStatus status)
    {
        if (!status.IsUsable || status.Manifest is not { } manifest)
        {
            throw new InvalidOperationException(status.Message);
        }

        var assemblyPath = Path.Combine(status.Directory, manifest.EngineAssembly);
        var context = new OcrPackLoadContext(assemblyPath);
        var assembly = context.LoadFromAssemblyPath(assemblyPath);

        var type = assembly.GetType(manifest.EngineType, throwOnError: false)
            ?? throw new InvalidOperationException(
                $"OCR Pack の中に {manifest.EngineType} が見つかりません。");

        return Activator.CreateInstance(type) as IOcrEngine
            ?? throw new InvalidOperationException("OCR Pack の形式が想定と違います。");
    }

    /// <summary>
    /// Pack の中の DLL を、Pack のフォルダーから解決する。
    /// ただし製品本体と共有する型(ExcelBatchTool.Core)は本体側から読む。
    /// そうしないと同じ名前の別の型になって、インターフェースが噛み合わなくなる。
    /// </summary>
    private sealed class OcrPackLoadContext(string assemblyPath)
        : AssemblyLoadContext(isCollectible: false)
    {
        private readonly AssemblyDependencyResolver _resolver = new(assemblyPath);

        protected override Assembly? Load(AssemblyName assemblyName)
        {
            if (assemblyName.Name is null
                || assemblyName.Name.StartsWith("ExcelBatchTool.Core", StringComparison.Ordinal))
            {
                return null;
            }

            var path = _resolver.ResolveAssemblyToPath(assemblyName);
            return path is null ? null : LoadFromAssemblyPath(path);
        }

        protected override nint LoadUnmanagedDll(string unmanagedDllName)
        {
            var path = _resolver.ResolveUnmanagedDllToPath(unmanagedDllName);
            return path is null ? nint.Zero : LoadUnmanagedDllFromPath(path);
        }
    }
}
