using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Path = System.IO.Path;
using Rectangle = System.Windows.Shapes.Rectangle;
using Size = System.Windows.Size;

namespace PptxViewer;

public record SlideData(int Index, string Notes, List<ShapeData> Shapes, string? BackgroundColor);

public record ShapeData(
    double X, double Y, double Width, double Height,
    string? Text, string? TextColor, double FontSize,
    bool IsBold, bool IsItalic, string? FillColor,
    bool IsTitle, TextAlignment TextAlign,
    string? ImagePath, bool HasBackground);

public static class PptxParser
{
    // EMU → WPF pixels  (1 inch = 914400 EMU, 96 DPI)
    private const double EmuToPx = 96.0 / 914400.0;

    public static List<SlideData> Parse(string filePath)
    {
        var slides = new List<SlideData>();

        using var prs = PresentationDocument.Open(filePath, false);
        var presentation = prs.PresentationPart?.Presentation
            ?? throw new InvalidOperationException("No presentation found");

        var slideIds = presentation.SlideIdList?.Elements<SlideId>().ToList()
                       ?? new List<SlideId>();

        int idx = 1;
        foreach (var slideId in slideIds)
        {
            var relId = slideId.RelationshipId?.Value;
            if (relId == null) continue;

            var slidePart = (SlidePart)prs.PresentationPart!.GetPartById(relId);
            slides.Add(ParseSlide(slidePart, idx++, prs));
        }

        return slides;
    }

    private static SlideData ParseSlide(SlidePart slidePart, int index, PresentationDocument prs)
    {
        var shapes = new List<ShapeData>();
        var notes = ExtractNotes(slidePart);
        var bgColor = ExtractBackground(slidePart, prs);

        var spTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (spTree == null) return new SlideData(index, notes, shapes, bgColor);

        // Regular shapes
        foreach (var sp in spTree.Elements<P.Shape>())
        {
            var shape = ParseShape(sp, slidePart);
            if (shape != null) shapes.Add(shape);
        }

        // Picture shapes
        foreach (var pic in spTree.Elements<P.Picture>())
        {
            var shape = ParsePicture(pic, slidePart);
            if (shape != null) shapes.Add(shape);
        }

        return new SlideData(index, notes, shapes, bgColor);
    }

    private static ShapeData? ParseShape(P.Shape sp, SlidePart slidePart)
    {
        var spPr = sp.ShapeProperties;
        var xfrm = spPr?.Transform2D;
        if (xfrm == null) return null;

        double x = (xfrm.Offset?.X?.Value ?? 0) * EmuToPx;
        double y = (xfrm.Offset?.Y?.Value ?? 0) * EmuToPx;
        double w = (xfrm.Extents?.Cx?.Value ?? 0) * EmuToPx;
        double h = (xfrm.Extents?.Cy?.Value ?? 0) * EmuToPx;

        // Fill color
        string? fillColor = null;
        var solidFill = spPr?.GetFirstChild<D.ShapeStyle>() == null
            ? spPr?.GetFirstChild<D.SolidFill>()
            : null;
        var spFill = spPr?.Elements<D.SolidFill>().FirstOrDefault();
        if (spFill?.RgbColorModelHex != null)
            fillColor = "#" + spFill.RgbColorModelHex.Val?.Value;

        // Text
        string? text = null;
        string? textColor = null;
        double fontSize = 18;
        bool isBold = false, isItalic = false;
        bool isTitle = false;
        var textAlign = TextAlignment.Left;

        var txBody = sp.TextBody;
        if (txBody != null)
        {
            var texts = new List<string>();
            foreach (var para in txBody.Elements<D.Paragraph>())
            {
                var paraText = string.Concat(para.Elements<D.Run>().Select(r => r.Text?.Text ?? ""));
                texts.Add(paraText);

                // Get formatting from first run
                var firstRun = para.Elements<D.Run>().FirstOrDefault();
                var rPr = firstRun?.RunProperties;
                if (rPr != null)
                {
                    if (rPr.FontSize?.Value is int fs && fs > 0)
                        fontSize = fs / 100.0;
                    isBold = rPr.Bold?.Value == true;
                    isItalic = rPr.Italic?.Value == true;

                    var colorElem = rPr.GetFirstChild<D.SolidFill>()?.RgbColorModelHex;
                    if (colorElem != null)
                        textColor = "#" + colorElem.Val?.Value;
                }

                // Paragraph alignment
                var pPr = para.ParagraphProperties;
                if (pPr?.Alignment?.Value == D.TextAlignmentTypeValues.Center)
                    textAlign = TextAlignment.Center;
                else if (pPr?.Alignment?.Value == D.TextAlignmentTypeValues.Right)
                    textAlign = TextAlignment.Right;
            }
            text = string.Join("\n", texts);
        }

        // Check if title placeholder
        var ph = sp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<P.PlaceholderShape>();
        if (ph?.Type?.Value == P.PlaceholderValues.Title ||
            ph?.Type?.Value == P.PlaceholderValues.CenteredTitle)
            isTitle = true;

        return new ShapeData(x, y, w, h, text, textColor, fontSize,
            isBold, isItalic, fillColor, isTitle, textAlign, null, false);
    }

    private static ShapeData? ParsePicture(P.Picture pic, SlidePart slidePart)
    {
        var spPr = pic.ShapeProperties;
        var xfrm = spPr?.Transform2D;
        if (xfrm == null) return null;

        double x = (xfrm.Offset?.X?.Value ?? 0) * EmuToPx;
        double y = (xfrm.Offset?.Y?.Value ?? 0) * EmuToPx;
        double w = (xfrm.Extents?.Cx?.Value ?? 0) * EmuToPx;
        double h = (xfrm.Extents?.Cy?.Value ?? 0) * EmuToPx;

        // Extract embedded image
        var blipFill = pic.BlipFill;
        var blip = blipFill?.Blip;
        var relId = blip?.Embed?.Value;
        if (relId == null) return null;

        try
        {
            var imgPart = (ImagePart)slidePart.GetPartById(relId);
            // Save to temp file
            var tmpFile = Path.Combine(Path.GetTempPath(), $"pptx_img_{Guid.NewGuid()}.png");
            using var stream = imgPart.GetStream();
            using var fs = File.Create(tmpFile);
            stream.CopyTo(fs);
            return new ShapeData(x, y, w, h, null, null, 0, false, false,
                null, false, TextAlignment.Left, tmpFile, false);
        }
        catch { return null; }
    }

    private static string ExtractNotes(SlidePart slidePart)
    {
        var notesPart = slidePart.NotesSlidePart;
        if (notesPart == null) return string.Empty;

        var texts = notesPart.NotesSlide.CommonSlideData?.ShapeTree?
            .Elements<P.Shape>()
            .Where(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>()?.Type?.Value == P.PlaceholderValues.Body)
            .SelectMany(s => s.TextBody?.Elements<D.Paragraph>() ?? [])
            .Select(p => string.Concat(p.Elements<D.Run>().Select(r => r.Text?.Text ?? "")))
            .Where(t => !string.IsNullOrWhiteSpace(t));

        return string.Join("\n", texts ?? []);
    }

    private static string? ExtractBackground(SlidePart slidePart, PresentationDocument prs)
    {
        // Try slide background first
        var bg = slidePart.Slide.CommonSlideData?.Background;
        var solidFill = bg?.BackgroundProperties?.GetFirstChild<D.SolidFill>();
        if (solidFill?.RgbColorModelHex?.Val?.Value is string hex)
            return "#" + hex;

        return null;
    }

    // ─── Render a slide onto a WPF Canvas ───
    public static void RenderSlide(SlideData slide, Canvas canvas)
    {
        canvas.Children.Clear();

        // Background
        canvas.Background = slide.BackgroundColor != null
            ? new SolidColorBrush(ParseColor(slide.BackgroundColor))
            : Brushes.White;

        foreach (var shape in slide.Shapes)
        {
            if (shape.ImagePath != null)
            {
                RenderImage(shape, canvas);
            }
            else if (!string.IsNullOrEmpty(shape.Text))
            {
                RenderTextShape(shape, canvas);
            }
            else if (shape.FillColor != null)
            {
                RenderRect(shape, canvas);
            }
        }
    }

    private static void RenderTextShape(ShapeData s, Canvas canvas)
    {
        // Optional fill rect behind text
        if (s.FillColor != null)
        {
            var rect = new Rectangle
            {
                Width = s.Width, Height = s.Height,
                Fill = new SolidColorBrush(ParseColor(s.FillColor))
            };
            Canvas.SetLeft(rect, s.X);
            Canvas.SetTop(rect, s.Y);
            canvas.Children.Add(rect);
        }

        var tb = new TextBlock
        {
            Text = s.Text,
            Width = s.Width,
            TextWrapping = TextWrapping.Wrap,
            Foreground = s.TextColor != null
                ? new SolidColorBrush(ParseColor(s.TextColor))
                : Brushes.Black,
            FontSize = Math.Max(s.FontSize, 10),
            FontWeight = s.IsBold ? FontWeights.Bold : FontWeights.Normal,
            FontStyle = s.IsItalic ? FontStyles.Italic : FontStyles.Normal,
            TextAlignment = s.TextAlign,
        };

        if (s.IsTitle)
        {
            tb.FontWeight = FontWeights.Bold;
            tb.FontSize = Math.Max(s.FontSize, 24);
        }

        Canvas.SetLeft(tb, s.X);
        Canvas.SetTop(tb, s.Y);
        canvas.Children.Add(tb);
    }

    private static void RenderRect(ShapeData s, Canvas canvas)
    {
        var rect = new Rectangle
        {
            Width = s.Width, Height = s.Height,
            Fill = new SolidColorBrush(ParseColor(s.FillColor!))
        };
        Canvas.SetLeft(rect, s.X);
        Canvas.SetTop(rect, s.Y);
        canvas.Children.Add(rect);
    }

    private static void RenderImage(ShapeData s, Canvas canvas)
    {
        try
        {
            var bmp = new BitmapImage();
            bmp.BeginInit();
            bmp.UriSource = new Uri(s.ImagePath!);
            bmp.CacheOption = BitmapCacheOption.OnLoad;
            bmp.EndInit();
            bmp.Freeze();

            var img = new System.Windows.Controls.Image
            {
                Source = bmp,
                Width = s.Width,
                Height = s.Height,
                Stretch = System.Windows.Media.Stretch.Uniform
            };
            Canvas.SetLeft(img, s.X);
            Canvas.SetTop(img, s.Y);
            canvas.Children.Add(img);
        }
        catch { /* skip broken images */ }
    }

    // ─── Thumbnail generation ───
    public static BitmapSource RenderThumbnail(SlideData slide, int width = 320, int height = 180)
    {
        var canvas = new Canvas { Width = 960, Height = 540 };
        RenderSlide(slide, canvas);
        canvas.Measure(new Size(960, 540));
        canvas.Arrange(new Rect(0, 0, 960, 540));

        var dpi = 96.0;
        var rtb = new RenderTargetBitmap(960, 540, dpi, dpi, PixelFormats.Pbgra32);
        rtb.Render(canvas);

        // Scale down
        var scaled = new TransformedBitmap(rtb,
            new ScaleTransform(width / 960.0, height / 540.0));
        scaled.Freeze();
        return scaled;
    }

    public static Color ParseColor(string hex)
    {
        try
        {
            hex = hex.TrimStart('#');
            if (hex.Length == 6)
            {
                return Color.FromRgb(
                    Convert.ToByte(hex[..2], 16),
                    Convert.ToByte(hex[2..4], 16),
                    Convert.ToByte(hex[4..6], 16));
            }
            if (hex.Length == 8)
            {
                return Color.FromArgb(
                    Convert.ToByte(hex[..2], 16),
                    Convert.ToByte(hex[2..4], 16),
                    Convert.ToByte(hex[4..6], 16),
                    Convert.ToByte(hex[6..8], 16));
            }
        }
        catch { }
        return Colors.Transparent;
    }
}
