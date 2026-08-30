using ExcelBatchTool.Core.Ocr;
using OpenCvSharp;

namespace ExcelBatchTool.Ocr;

/// <summary>ページ画像から測る・直す。OCR にかける前の下ごしらえ。</summary>
internal static class PageImageOps
{
    /// <summary>罫線とみなすのに必要な長さ(ページの幅・高さに対する割合)。</summary>
    private const double RulingLengthRatio = 0.3;

    /// <summary>黒とみなす明るさ。</summary>
    private const int DarkLevel = 128;

    public static Mat Binarize(Mat source)
    {
        using var gray = source.CvtColor(ColorConversionCodes.BGR2GRAY);
        var binary = new Mat();
        Cv2.Threshold(gray, binary, 0, 255, ThresholdTypes.BinaryInv | ThresholdTypes.Otsu);
        return binary;
    }

    /// <summary>
    /// 傾きの推定。文字を横につないで行の塊にし、その塊の傾きの中央値を取る。
    /// 行らしい塊が少ないときは信頼できないものとして扱う(勝手に回さない)。
    ///
    /// 罫線は先に取り除く。表のページでは罫線が画面を横切る 1 本の巨大な塊になり、
    /// そのままだと文字の行より強く効いて角度がまるで当てにならなくなる
    /// (実測: 傾き 0 度の罫線表が「6 度超」と判定されて止まった)。
    /// </summary>
    /// <summary>
    /// 罫線そのものから傾きを測る。
    ///
    /// 罫線のある表では、こちらのほうが文字の塊より正確に測れる。罫線はページを
    /// 横切るほど長く、途切れも少ないため、角度がほとんどぶれない。
    ///
    /// なぜ要るか: 文字から測った角度で直しても、わずかな残り傾き
    /// (平均 0.33 度)が残る。表ではこれが致命的で、ページの端へ行くほど
    /// 行がずれ、**セルの中身が 1 行ずつずれた表**が出来上がる。実測では
    /// 2 度傾いた罫線表でセル一致 19.0%、誤って自動確定したセルが 112 件だった。
    /// しかも中身は埋まっていて列の形も揃うので、どの安全弁にも掛からない。
    /// </summary>
    public static (double Degrees, bool Reliable) SkewFromRulings(Mat forRulings)
    {
        // 直線そのものを探す。横向きの型で抜く方法は使えない
        //  ― 2 度傾いただけで、幅 2000 画素の罫線は端で 70 画素も下がるため、
        //    まっすぐな型には一切引っかからない(実測で罫線 0 本になっていた)。
        var minimumLength = Math.Max(forRulings.Width / 3, 60);
        using var segments = new Mat();
        var lines = Cv2.HoughLinesP(
            forRulings,
            rho: 1,
            theta: Math.PI / 1800,
            threshold: minimumLength / 2,
            minLineLength: minimumLength,
            maxLineGap: 20);

        var angles = new List<double>();
        foreach (var line in lines)
        {
            var dx = (double)(line.P2.X - line.P1.X);
            var dy = (double)(line.P2.Y - line.P1.Y);
            if (Math.Abs(dx) < 1)
            {
                continue;
            }

            var angle = Math.Atan2(dy, dx) * 180 / Math.PI;

            // 横罫線だけを見る(縦罫線と枠の縁は除く)。
            if (Math.Abs(angle) > 20)
            {
                continue;
            }

            angles.Add(angle);
        }

        if (angles.Count < 3)
        {
            return (0, false);
        }

        angles.Sort();
        return (angles[angles.Count / 2], true);
    }

    public static (double Degrees, bool Reliable) Skew(Mat binary)
    {
        using var withoutRulings = RemoveRulings(binary);
        using var kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(25, 3));
        using var merged = new Mat();
        Cv2.MorphologyEx(withoutRulings, merged, MorphTypes.Close, kernel);

        Cv2.FindContours(
            merged, out var contours, out _,
            RetrievalModes.External, ContourApproximationModes.ApproxSimple);

        var angles = new List<double>();
        foreach (var contour in contours)
        {
            var rect = Cv2.MinAreaRect(contour);
            var width = Math.Max(rect.Size.Width, rect.Size.Height);
            var height = Math.Min(rect.Size.Width, rect.Size.Height);

            // 明らかに行らしい塊(横に長い)だけを見る。
            if (width < 60 || height < 3 || width < height * 4)
            {
                continue;
            }

            var angle = rect.Angle;
            if (rect.Size.Width < rect.Size.Height)
            {
                angle += 90;
            }

            while (angle > 45)
            {
                angle -= 90;
            }

            while (angle < -45)
            {
                angle += 90;
            }

            angles.Add(angle);
        }

        if (angles.Count < 5)
        {
            return (0, false);
        }

        angles.Sort();
        var median = angles[angles.Count / 2];

        // 塊ごとの角度がばらついているなら、その推定は当てにならない。
        var spread = angles[(int)(angles.Count * 0.75)] - angles[(int)(angles.Count * 0.25)];
        return (median, spread <= 2.0);
    }

    /// <summary>罫線を消した画像(文字だけを残す)。</summary>
    private static Mat RemoveRulings(Mat binary)
    {
        using var horizontalKernel = Cv2.GetStructuringElement(
            MorphShapes.Rect, new Size(Math.Max(20, (int)(binary.Width * 0.15)), 1));
        using var horizontal = binary.MorphologyEx(MorphTypes.Open, horizontalKernel);

        using var verticalKernel = Cv2.GetStructuringElement(
            MorphShapes.Rect, new Size(1, Math.Max(20, (int)(binary.Height * 0.15))));
        using var vertical = binary.MorphologyEx(MorphTypes.Open, verticalKernel);

        using var rulings = new Mat();
        Cv2.BitwiseOr(horizontal, vertical, rulings);

        // 罫線を少し太らせてから引く(線の縁が残って点々にならないように)。
        using var thicken = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(3, 3));
        using var thick = rulings.Dilate(thicken);

        var result = new Mat();
        Cv2.Subtract(binary, thick, result);
        return result;
    }

    /// <summary>画像を回してまっすぐにする。端が切れないよう外接する大きさへ広げる。</summary>
    public static Mat Rotate(Mat source, double degrees, out DeskewTransform transform)
    {
        var centerX = source.Width / 2.0;
        var centerY = source.Height / 2.0;

        // 向きは DeskewTransform.FromRotation が決める。ここで符号を触らない
        // (触ったせいで確認画面の枠が文字からずれていた)。
        transform = DeskewTransform.FromRotation(degrees, centerX, centerY);

        using var matrix = Cv2.GetRotationMatrix2D(
            new Point2f((float)centerX, (float)centerY), degrees, 1);

        var rotated = new Mat();
        Cv2.WarpAffine(
            source, rotated, matrix, source.Size(),
            InterpolationFlags.Cubic, BorderTypes.Replicate);

        return rotated;
    }

    /// <summary>
    /// 罫線を拾うための二値化。
    ///
    /// ページ全体の明るさで一律に切る(大津)と、細い罫線が落ちる。実測では、
    /// 表の外枠(いちばん外側の縦線 2 本)だけが消えて列が 4 → 2 になり、
    /// セルの中身が 1 列ずつずれた。周りとの差で切る方法なら細い線も残る。
    /// </summary>
    public static Mat BinarizeForRulings(Mat source)
    {
        using var gray = source.CvtColor(ColorConversionCodes.BGR2GRAY);
        var binary = new Mat();
        Cv2.AdaptiveThreshold(
            gray, binary, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.BinaryInv, 25, 15);
        return binary;
    }

    /// <summary>罫線(長い横線・縦線)の位置。</summary>
    public static (List<double> Rows, List<double> Columns) Rulings(Mat binary)
    {
        using var horizontalKernel = Cv2.GetStructuringElement(
            MorphShapes.Rect, new Size(Math.Max(20, (int)(binary.Width * RulingLengthRatio)), 1));
        using var horizontal = binary.MorphologyEx(MorphTypes.Open, horizontalKernel);

        using var verticalKernel = Cv2.GetStructuringElement(
            MorphShapes.Rect, new Size(1, Math.Max(20, (int)(binary.Height * RulingLengthRatio))));
        using var vertical = binary.MorphologyEx(MorphTypes.Open, verticalKernel);

        return (Positions(horizontal, alongRows: true), Positions(vertical, alongRows: false));
    }

    /// <summary>
    /// 記入欄の下線の本数。短めの横線で、その下に縦線が来ないものを数える。
    /// 表の罫線と区別するために、縦線と組にならないことを条件にする。
    /// </summary>
    public static int Underlines(Mat binary)
    {
        using var kernel = Cv2.GetStructuringElement(
            MorphShapes.Rect, new Size(Math.Max(20, (int)(binary.Width * 0.08)), 1));
        using var lines = binary.MorphologyEx(MorphTypes.Open, kernel);

        Cv2.FindContours(
            lines, out var contours, out _,
            RetrievalModes.External, ContourApproximationModes.ApproxSimple);

        var count = 0;
        foreach (var contour in contours)
        {
            var rect = Cv2.BoundingRect(contour);

            // ページを横切るような長い線は罫線。記入欄の下線はもっと短い。
            if (rect.Height <= 4
                && rect.Width >= binary.Width * 0.08
                && rect.Width <= binary.Width * RulingLengthRatio)
            {
                count++;
            }
        }

        return count;
    }

    /// <summary>指定した場所の黒い画素の割合。</summary>
    public static double InkRatio(Mat gray, OcrBox area)
    {
        var rect = new Rect(
            (int)Math.Round(area.X),
            (int)Math.Round(area.Y),
            (int)Math.Round(area.Width),
            (int)Math.Round(area.Height));

        rect = rect.Intersect(new Rect(0, 0, gray.Width, gray.Height));
        if (rect.Width <= 0 || rect.Height <= 0)
        {
            return 0;
        }

        using var region = new Mat(gray, rect);
        using var dark = new Mat();
        Cv2.Threshold(region, dark, DarkLevel, 255, ThresholdTypes.BinaryInv);

        return Cv2.CountNonZero(dark) / (double)(rect.Width * rect.Height);
    }

    /// <summary>線画像を行(列)方向に足し込み、山の位置を線の座標として拾う。</summary>
    private static List<double> Positions(Mat lines, bool alongRows)
    {
        using var reduced = new Mat();
        Cv2.Reduce(
            lines, reduced,
            alongRows ? ReduceDimension.Column : ReduceDimension.Row,
            ReduceTypes.Avg, MatType.CV_32F);

        var length = alongRows ? lines.Rows : lines.Cols;
        var values = new float[length];
        for (var index = 0; index < length; index++)
        {
            values[index] = alongRows ? reduced.At<float>(index, 0) : reduced.At<float>(0, index);
        }

        var positions = new List<double>();
        const float threshold = 40f;
        var start = -1;

        for (var index = 0; index < length; index++)
        {
            if (values[index] >= threshold)
            {
                if (start < 0)
                {
                    start = index;
                }
            }
            else if (start >= 0)
            {
                positions.Add((start + index - 1) / 2.0);
                start = -1;
            }
        }

        if (start >= 0)
        {
            positions.Add((start + length - 1) / 2.0);
        }

        return positions;
    }
}
