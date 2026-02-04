using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using WebPptx.Controllers;
using A = DocumentFormat.OpenXml.Drawing;

namespace WebPptx.Services;

public interface IHtmlPageService
{
    HtmlPageResponse Export(HtmlPageRequest request);
}

public class HtmlPageService : IHtmlPageService
{
    private readonly IConfiguration _configuration;
    private readonly IHostEnvironment _environment;

    public HtmlPageService(IConfiguration configuration, IHostEnvironment environment)
    {
        _configuration = configuration;
        _environment = environment;
    }

    public HtmlPageResponse Export(HtmlPageRequest request)
    {
        if (request is null)
        {
            throw new ArgumentException("Body is required.");
        }

        if (string.IsNullOrWhiteSpace(request.Path))
        {
            throw new ArgumentException("Path must be a non-empty string.");
        }

        var fullPath = Path.GetFullPath(request.Path);
        if (!System.IO.File.Exists(fullPath))
        {
            throw new FileNotFoundException("File not found.", fullPath);
        }

        if (!string.Equals(Path.GetExtension(fullPath), ".pptx", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Only .pptx files are supported.");
        }

        var options = BuildOptions(request);

        var samplesRoot = GetSamplesRoot();
        Directory.CreateDirectory(samplesRoot);

        var outputDir = Path.Combine(samplesRoot, Path.GetFileNameWithoutExtension(fullPath));
        Directory.CreateDirectory(outputDir);

        var pagesDir = Path.Combine(outputDir, "htmlpage");
        EnsureEmptyDirectory(pagesDir);

        var pdfPath = ConvertPptxToPdf(fullPath, outputDir, options.SofficePath);

        var attachments = new List<HtmlAttachment>();
        var slideModels = new List<HtmlSlide>();
        var imageAssets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        SlideSizeInfo slideSizeInfo;
        var slideCount = 0;

        var attachmentsDir = Path.Combine(pagesDir, "attachments");
        var imagesDir = Path.Combine(pagesDir, "images");
        Directory.CreateDirectory(attachmentsDir);
        Directory.CreateDirectory(imagesDir);

        using (var presentation = PresentationDocument.Open(fullPath, false))
        {
            var presentationPart = presentation.PresentationPart;
            if (presentationPart is null)
            {
                throw new InvalidOperationException("Presentation part not found.");
            }

            slideSizeInfo = GetSlideSizeInfo(presentationPart);
            var colorScheme = presentationPart.ThemePart?.Theme?.ThemeElements?.ColorScheme;
            var slideIds = presentationPart.Presentation?.SlideIdList?.ChildElements.OfType<SlideId>().ToList()
                ?? new List<SlideId>();
            slideCount = slideIds.Count;

            var slideIndex = 0;
            foreach (var slideId in slideIds)
            {
                slideIndex++;
                var relationshipId = slideId.RelationshipId?.Value;
                if (string.IsNullOrWhiteSpace(relationshipId))
                {
                    continue;
                }

                if (presentationPart.GetPartById(relationshipId) is not SlidePart slidePart)
                {
                    continue;
                }

                var textBlocks = ExtractTextBlocks(slidePart, colorScheme);
                var images = ExtractImageBlocks(slidePart, imagesDir, imageAssets);
                var backgroundColor = GetSlideBackgroundColor(slidePart, colorScheme);
                var slideAttachments = ExtractAttachments(slidePart, attachmentsDir, slideIndex);
                attachments.AddRange(slideAttachments);
                slideModels.Add(new HtmlSlide(slideIndex, textBlocks, images, backgroundColor));
            }
        }

        var logoHashes = slideModels
            .SelectMany(slide => slide.Images.Select(image => new { slide.Index, Image = image }))
            .Where(item => IsLogoCandidate(item.Image, slideSizeInfo))
            .GroupBy(item => item.Image.Hash, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Select(item => item.Index).Distinct().Count() > 1)
            .Select(group => group.Key)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var logos = slideModels
            .SelectMany(slide => slide.Images)
            .Where(image => logoHashes.Contains(image.Hash))
            .GroupBy(image => image.Hash, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToList();

        var filteredSlides = slideModels
            .Select(slide => new HtmlSlide(
                slide.Index,
                slide.TextBlocks,
                slide.Images.Where(image => !logoHashes.Contains(image.Hash)).ToList(),
                slide.BackgroundColor))
            .ToList();

        var htmlSourcePath = ConvertPptxToHtml(fullPath, pagesDir, options.SofficePath);
        var htmlSource = System.IO.File.ReadAllText(htmlSourcePath, Encoding.UTF8);

        var htmlPath = Path.Combine(pagesDir, "index.html");
        var title = Path.GetFileNameWithoutExtension(fullPath);
        var html = BuildHtmlFromLibreOffice(
            htmlSource,
            title,
            pagesDir,
            slideSizeInfo,
            filteredSlides,
            logos,
            attachments);
        System.IO.File.WriteAllText(htmlPath, html, Encoding.UTF8);

        var imageFiles = imageAssets.Values.ToList();
        return new HtmlPageResponse(fullPath, outputDir, htmlPath, pdfPath, slideCount, imageFiles.Count, imageFiles);
    }

    private HtmlPageOptions BuildOptions(HtmlPageRequest request)
    {
        var maxWidth = request.MaxWidth
            ?? _configuration.GetValue<int?>("Screenshots:MaxWidth");
        var maxHeight = request.MaxHeight
            ?? _configuration.GetValue<int?>("Screenshots:MaxHeight");
        var jpegQuality = request.JpegQuality
            ?? _configuration.GetValue("Screenshots:JpegQuality", 70);
        jpegQuality = Math.Clamp(jpegQuality, 1, 100);

        var pdfDpi = request.PdfDpi
            ?? _configuration.GetValue<int?>("Screenshots:PdfDpi");
        if (pdfDpi is not null && pdfDpi <= 0)
        {
            pdfDpi = null;
        }
        else if (pdfDpi is not null)
        {
            pdfDpi = Math.Clamp(pdfDpi.Value, 72, 300);
        }

        var configuredSofficePath = string.IsNullOrWhiteSpace(request.SofficePath)
            ? _configuration["LibreOffice:SofficePath"]
            : request.SofficePath;

        return new HtmlPageOptions(
            configuredSofficePath,
            maxWidth,
            maxHeight,
            jpegQuality,
            pdfDpi);
    }

    private string GetSamplesRoot()
    {
        var contentRoot = _environment.ContentRootPath;
        var parent = Directory.GetParent(contentRoot)?.FullName;
        var root = string.IsNullOrWhiteSpace(parent) ? contentRoot : parent;
        return Path.Combine(root, "samples");
    }

    private static void EnsureEmptyDirectory(string directory)
    {
        if (Directory.Exists(directory))
        {
            foreach (var file in Directory.EnumerateFiles(directory))
            {
                System.IO.File.Delete(file);
            }

            foreach (var dir in Directory.EnumerateDirectories(directory))
            {
                Directory.Delete(dir, recursive: true);
            }

            return;
        }

        Directory.CreateDirectory(directory);
    }

    private static SlideSizeInfo GetSlideSizeInfo(PresentationPart presentationPart)
    {
        var slideSize = presentationPart.Presentation?.SlideSize;
        if (slideSize is null)
        {
            return new SlideSizeInfo(0, 0);
        }

        var width = GetLong(slideSize.Cx) ?? 0;
        var height = GetLong(slideSize.Cy) ?? 0;
        return new SlideSizeInfo(width, height);
    }

    private static string? GetSlideBackgroundColor(SlidePart slidePart, A.ColorScheme? colorScheme)
    {
        var background = slidePart.Slide.CommonSlideData?.Background;
        var properties = background?.BackgroundProperties;
        var solidFill = properties?.GetFirstChild<A.SolidFill>();
        if (solidFill is null)
        {
            return null;
        }

        if (solidFill.RgbColorModelHex is not null)
        {
            var hex = GetStringValue(solidFill.RgbColorModelHex.Val);
            return string.IsNullOrWhiteSpace(hex) ? null : $"#{hex}";
        }

        if (solidFill.SchemeColor is not null && colorScheme is not null)
        {
            var hex = ResolveSchemeColor(solidFill.SchemeColor.Val, colorScheme);
            return string.IsNullOrWhiteSpace(hex) ? null : $"#{hex}";
        }

        return null;
    }

    private string ConvertPptxToPdf(string pptxPath, string outputDir, string? configuredSofficePath)
    {
        var sofficePath = PptxController.ResolveSofficePath(configuredSofficePath);
        if (sofficePath is null)
        {
            throw new InvalidOperationException("LibreOffice (soffice) not found. Provide 'sofficePath' or set LibreOffice:SofficePath.");
        }

        var args = $"--headless --nologo --nolockcheck --norestore --convert-to pdf --outdir \"{outputDir}\" \"{pptxPath}\"";
        var startInfo = new ProcessStartInfo
        {
            FileName = sofficePath,
            Arguments = args,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var process = Process.Start(startInfo);
        if (process is null)
        {
            throw new InvalidOperationException("Failed to start LibreOffice.");
        }

        _ = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();

        if (!process.WaitForExit(TimeSpan.FromSeconds(120)))
        {
            try
            {
                process.Kill(true);
            }
            catch
            {
            }

            throw new InvalidOperationException("LibreOffice PDF export timed out.");
        }

        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException($"LibreOffice PDF export failed: {stderr}");
        }

        var baseName = Path.GetFileNameWithoutExtension(pptxPath);
        var expectedPdf = Path.Combine(outputDir, $"{baseName}.pdf");
        if (System.IO.File.Exists(expectedPdf))
        {
            return expectedPdf;
        }

        var pdfPath = Directory.EnumerateFiles(outputDir, "*.pdf", SearchOption.TopDirectoryOnly)
            .FirstOrDefault();

        if (string.IsNullOrWhiteSpace(pdfPath))
        {
            throw new InvalidOperationException("LibreOffice did not produce a PDF file.");
        }

        return pdfPath;
    }

    private string ConvertPptxToHtml(string pptxPath, string outputDir, string? configuredSofficePath)
    {
        var sofficePath = PptxController.ResolveSofficePath(configuredSofficePath);
        if (sofficePath is null)
        {
            throw new InvalidOperationException("LibreOffice (soffice) not found. Provide 'sofficePath' or set LibreOffice:SofficePath.");
        }

        var args = $"--headless --nologo --nolockcheck --norestore --convert-to \"html:impress_html_Export\" --outdir \"{outputDir}\" \"{pptxPath}\"";
        var startInfo = new ProcessStartInfo
        {
            FileName = sofficePath,
            Arguments = args,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var process = Process.Start(startInfo);
        if (process is null)
        {
            throw new InvalidOperationException("Failed to start LibreOffice.");
        }

        _ = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();

        if (!process.WaitForExit(TimeSpan.FromSeconds(120)))
        {
            try
            {
                process.Kill(true);
            }
            catch
            {
            }

            throw new InvalidOperationException("LibreOffice HTML export timed out.");
        }

        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException($"LibreOffice HTML export failed: {stderr}");
        }

        var htmlPath = Directory.EnumerateFiles(outputDir, "*.*", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(path =>
                string.Equals(Path.GetExtension(path), ".html", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(Path.GetExtension(path), ".htm", StringComparison.OrdinalIgnoreCase));

        if (string.IsNullOrWhiteSpace(htmlPath))
        {
            throw new InvalidOperationException("LibreOffice did not produce an HTML file.");
        }

        return htmlPath;
    }

    private sealed record HtmlSlide(int Index, List<HtmlTextBlock> TextBlocks, List<HtmlImageBlock> Images, string? BackgroundColor);

    private sealed record HtmlTextBlock(long X, long Y, long Cx, long Cy, HtmlPadding Padding, string Html);

    private sealed record HtmlImageBlock(long X, long Y, long Cx, long Cy, string FilePath, string Hash);

    private sealed record HtmlAttachment(int SlideIndex, string FilePath);

    private readonly record struct HtmlPadding(double LeftPx, double TopPx, double RightPx, double BottomPx);

    private enum ListKind
    {
        None,
        Bullet,
        Numbered
    }

    private static List<HtmlTextBlock> ExtractTextBlocks(SlidePart slidePart, A.ColorScheme? colorScheme)
    {
        var blocks = new List<HtmlTextBlock>();

        foreach (var shape in slidePart.Slide.Descendants<Shape>())
        {
            if (shape.TextBody is null)
            {
                continue;
            }

            if (!TryGetTransform(shape.ShapeProperties?.Transform2D, out var rect))
            {
                continue;
            }

            var html = BuildTextBodyHtml(shape.TextBody, colorScheme);
            if (string.IsNullOrWhiteSpace(html))
            {
                continue;
            }

            var padding = GetBodyPadding(shape.TextBody.BodyProperties);
            blocks.Add(new HtmlTextBlock(rect.X, rect.Y, rect.Cx, rect.Cy, padding, html));
        }

        return blocks;
    }

    private static List<HtmlImageBlock> ExtractImageBlocks(
        SlidePart slidePart,
        string imagesDir,
        Dictionary<string, string> imageAssets)
    {
        var blocks = new List<HtmlImageBlock>();

        foreach (var picture in slidePart.Slide.Descendants<Picture>())
        {
            if (!TryGetTransform(picture.ShapeProperties?.Transform2D, out var rect))
            {
                continue;
            }

            var blip = picture.BlipFill?.Blip;
            var relationshipId = blip?.Embed?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId))
            {
                continue;
            }

            if (slidePart.GetPartById(relationshipId) is not ImagePart imagePart)
            {
                continue;
            }

            string hash;
            using (var input = imagePart.GetStream())
            using (var buffer = new MemoryStream())
            {
                input.CopyTo(buffer);
                var bytes = buffer.ToArray();
                hash = ComputeSha256Hex(bytes);

                if (!imageAssets.TryGetValue(hash, out var storedPath))
                {
                    var extension = GetImageExtension(imagePart);
                    var token = hash.Length > 12 ? hash[..12] : hash;
                    var fileName = $"asset-{token}{extension}";
                    storedPath = Path.Combine(imagesDir, fileName);
                    System.IO.File.WriteAllBytes(storedPath, bytes);
                    imageAssets[hash] = storedPath;
                }
            }

            var filePath = imageAssets[hash];
            blocks.Add(new HtmlImageBlock(rect.X, rect.Y, rect.Cx, rect.Cy, filePath, hash));
        }

        return blocks;
    }

    private static List<HtmlAttachment> ExtractAttachments(SlidePart slidePart, string attachmentsDir, int slideIndex)
    {
        var attachments = new List<HtmlAttachment>();
        var attachmentIndex = 0;

        foreach (var part in slidePart.Parts)
        {
            if (part.OpenXmlPart is ImagePart)
            {
                continue;
            }

            if (part.OpenXmlPart is EmbeddedPackagePart or EmbeddedObjectPart)
            {
                attachmentIndex++;
                var extension = GetPartExtension(part.OpenXmlPart);
                var fileName = $"slide-{slideIndex:000}-attachment{attachmentIndex}{extension}";
                var destinationPath = Path.Combine(attachmentsDir, fileName);

                using var input = part.OpenXmlPart.GetStream(FileMode.Open, FileAccess.Read);
                using var output = System.IO.File.Create(destinationPath);
                input.CopyTo(output);

                attachments.Add(new HtmlAttachment(slideIndex, destinationPath));
            }
        }

        return attachments;
    }

    private static bool TryGetTransform(A.Transform2D? transform, out SlideFrameRect rect)
    {
        rect = default;
        if (transform?.Offset is null || transform.Extents is null)
        {
            return false;
        }

        var x = GetLong(transform.Offset?.X) ?? 0;
        var y = GetLong(transform.Offset?.Y) ?? 0;
        var cx = GetLong(transform.Extents?.Cx) ?? 0;
        var cy = GetLong(transform.Extents?.Cy) ?? 0;

        if (cx <= 0 || cy <= 0)
        {
            return false;
        }

        rect = new SlideFrameRect(x, y, cx, cy);
        return true;
    }

    private static HtmlPadding GetBodyPadding(A.BodyProperties? bodyProperties)
    {
        if (bodyProperties is null)
        {
            return new HtmlPadding(0, 0, 0, 0);
        }

        var left = GetLong(bodyProperties.LeftInset) ?? 0;
        var right = GetLong(bodyProperties.RightInset) ?? 0;
        var top = GetLong(bodyProperties.TopInset) ?? 0;
        var bottom = GetLong(bodyProperties.BottomInset) ?? 0;

        return new HtmlPadding(
            EmuToPx(left),
            EmuToPx(top),
            EmuToPx(right),
            EmuToPx(bottom));
    }

    private static string BuildTextBodyHtml(TextBody textBody, A.ColorScheme? colorScheme)
    {
        var sb = new StringBuilder();

        var inList = false;
        var currentListTag = "ul";

        foreach (var paragraph in textBody.Descendants<A.Paragraph>())
        {
            var paragraphProps = paragraph.ParagraphProperties;
            var listKind = GetListKind(paragraphProps);
            var listTag = listKind == ListKind.Numbered ? "ol" : "ul";

            if (listKind != ListKind.None)
            {
                if (!inList || !string.Equals(currentListTag, listTag, StringComparison.Ordinal))
                {
                    if (inList)
                    {
                        sb.Append($"</{currentListTag}>");
                    }

                    sb.Append($"<{listTag} class=\"pptx-list\">");
                    inList = true;
                    currentListTag = listTag;
                }

                var itemStyle = BuildParagraphStyle(paragraphProps, isListItem: true);
                if (string.IsNullOrWhiteSpace(itemStyle))
                {
                    sb.Append("<li class=\"pptx-list-item\">");
                }
                else
                {
                    sb.Append($"<li class=\"pptx-list-item\" style=\"{itemStyle}\">");
                }

                sb.Append(BuildParagraphRunsHtml(paragraph, GetDefaultRunProperties(paragraphProps), colorScheme));
                sb.Append("</li>");
            }
            else
            {
                if (inList)
                {
                    sb.Append($"</{currentListTag}>");
                    inList = false;
                }

                var paragraphStyle = BuildParagraphStyle(paragraphProps, isListItem: false);
                if (string.IsNullOrWhiteSpace(paragraphStyle))
                {
                    sb.Append("<div class=\"pptx-paragraph\">");
                }
                else
                {
                    sb.Append($"<div class=\"pptx-paragraph\" style=\"{paragraphStyle}\">");
                }

                sb.Append(BuildParagraphRunsHtml(paragraph, GetDefaultRunProperties(paragraphProps), colorScheme));
                sb.Append("</div>");
            }
        }

        if (inList)
        {
            sb.Append($"</{currentListTag}>");
        }

        return sb.ToString();
    }

    private static string BuildParagraphRunsHtml(A.Paragraph paragraph, A.DefaultRunProperties? defaultRunProperties, A.ColorScheme? colorScheme)
    {
        var sb = new StringBuilder();

        foreach (var child in paragraph.ChildElements)
        {
            switch (child)
            {
                case A.Run run:
                    sb.Append(BuildRunHtml(run, defaultRunProperties, colorScheme));
                    break;
                case A.Field field:
                    sb.Append(BuildFieldHtml(field, defaultRunProperties, colorScheme));
                    break;
                case A.Break:
                    sb.Append("<br>");
                    break;
            }
        }

        return sb.ToString();
    }

    private static ListKind GetListKind(A.ParagraphProperties? paragraphProperties)
    {
        if (paragraphProperties is null)
        {
            return ListKind.None;
        }

        if (paragraphProperties.GetFirstChild<A.NoBullet>() is not null)
        {
            return ListKind.None;
        }

        if (paragraphProperties.GetFirstChild<A.AutoNumberedBullet>() is not null)
        {
            return ListKind.Numbered;
        }

        if (paragraphProperties.GetFirstChild<A.CharacterBullet>() is not null ||
            paragraphProperties.GetFirstChild<A.PictureBullet>() is not null)
        {
            return ListKind.Bullet;
        }

        return ListKind.None;
    }

    private static A.DefaultRunProperties? GetDefaultRunProperties(A.ParagraphProperties? paragraphProperties)
    {
        return paragraphProperties?.GetFirstChild<A.DefaultRunProperties>();
    }

    private static string BuildParagraphStyle(A.ParagraphProperties? paragraphProperties, bool isListItem)
    {
        if (paragraphProperties is null)
        {
            return string.Empty;
        }

        var styles = new List<string>();
        var align = MapAlignment(paragraphProperties.Alignment);
        if (!string.IsNullOrWhiteSpace(align))
        {
            styles.Add($"text-align:{align}");
        }

        var lineHeight = BuildLineHeightStyle(paragraphProperties.LineSpacing);
        if (!string.IsNullOrWhiteSpace(lineHeight))
        {
            styles.Add(lineHeight);
        }

        var before = BuildSpacingStyle(paragraphProperties.SpaceBefore, "margin-top");
        if (!string.IsNullOrWhiteSpace(before))
        {
            styles.Add(before);
        }

        var after = BuildSpacingStyle(paragraphProperties.SpaceAfter, "margin-bottom");
        if (!string.IsNullOrWhiteSpace(after))
        {
            styles.Add(after);
        }

        var marginLeft = BuildEmuStyle(GetLong(paragraphProperties.LeftMargin), "margin-left");
        if (!string.IsNullOrWhiteSpace(marginLeft))
        {
            styles.Add(marginLeft);
        }

        var indent = BuildEmuStyle(GetLong(paragraphProperties.Indent), "text-indent");
        if (!string.IsNullOrWhiteSpace(indent))
        {
            styles.Add(indent);
        }

        if (string.IsNullOrWhiteSpace(marginLeft))
        {
            var levelIndent = BuildLevelIndentStyle(GetInt(paragraphProperties.Level), isListItem);
            if (!string.IsNullOrWhiteSpace(levelIndent))
            {
                styles.Add(levelIndent);
            }
        }

        return styles.Count == 0 ? string.Empty : string.Join(";", styles) + ";";
    }

    private static string? BuildLineHeightStyle(A.LineSpacing? lineSpacing)
    {
        if (lineSpacing is null)
        {
            return null;
        }

        var points = GetInt(lineSpacing.GetFirstChild<A.SpacingPoints>()?.Val);
        if (points is not null && points > 0)
        {
            var px = points.Value / 100.0 * 96.0 / 72.0;
            return $"line-height:calc(var(--scale,1) * {px:0.###}px)";
        }

        var percent = GetInt(lineSpacing.GetFirstChild<A.SpacingPercent>()?.Val);
        if (percent is not null && percent > 0)
        {
            var ratio = percent.Value / 100000.0;
            return $"line-height:{ratio:0.###}";
        }

        return null;
    }

    private static string? BuildSpacingStyle(OpenXmlElement? spacingElement, string cssProperty)
    {
        if (spacingElement is null)
        {
            return null;
        }

        var points = GetInt(spacingElement.GetFirstChild<A.SpacingPoints>()?.Val);
        if (points is not null && points > 0)
        {
            var px = points.Value / 100.0 * 96.0 / 72.0;
            return $"{cssProperty}:calc(var(--scale,1) * {px:0.###}px)";
        }

        var percent = GetInt(spacingElement.GetFirstChild<A.SpacingPercent>()?.Val);
        if (percent is not null && percent > 0)
        {
            var ratio = percent.Value / 100000.0;
            return $"{cssProperty}:{ratio:0.###}em";
        }

        return null;
    }

    private static string? BuildEmuStyle(long? emuValue, string cssProperty)
    {
        if (emuValue is null)
        {
            return null;
        }

        var px = EmuToPx(emuValue.Value);
        if (Math.Abs(px) < 0.001)
        {
            return null;
        }

        return $"{cssProperty}:calc(var(--scale,1) * {px:0.###}px)";
    }

    private static string? BuildLevelIndentStyle(int? level, bool isListItem)
    {
        if (level is null || level <= 0)
        {
            return null;
        }

        var baseIndentPx = 24.0;
        var px = level.Value * baseIndentPx;
        var property = isListItem ? "margin-left" : "margin-left";
        return $"{property}:calc(var(--scale,1) * {px:0.###}px)";
    }

    private static string? MapAlignment(OpenXmlSimpleType? alignment)
    {
        var token = NormalizeToken(GetStringValue(alignment));
        if (string.IsNullOrWhiteSpace(token))
        {
            return null;
        }
        if (token is "ctr" or "center" or "centre")
        {
            return "center";
        }
        if (token is "r" or "right")
        {
            return "right";
        }
        if (token.Contains("just", StringComparison.Ordinal) ||
            token.Contains("dist", StringComparison.Ordinal))
        {
            return "justify";
        }

        return "left";
    }

    private static string BuildRunHtml(A.Run run, A.DefaultRunProperties? defaultRunProperties, A.ColorScheme? colorScheme)
    {
        var text = string.Concat(run.Descendants<A.Text>().Select(item => item.Text));
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var encoded = WebUtility.HtmlEncode(text);
        var style = BuildRunStyle(run.RunProperties, defaultRunProperties, colorScheme);
        if (string.IsNullOrWhiteSpace(style))
        {
            return encoded;
        }

        return $"<span style=\"{style}\">{encoded}</span>";
    }

    private static string BuildFieldHtml(A.Field field, A.DefaultRunProperties? defaultRunProperties, A.ColorScheme? colorScheme)
    {
        var text = string.Concat(field.Descendants<A.Text>().Select(item => item.Text));
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var encoded = WebUtility.HtmlEncode(text);
        var style = BuildRunStyle(field.RunProperties, defaultRunProperties, colorScheme);
        if (string.IsNullOrWhiteSpace(style))
        {
            return encoded;
        }

        return $"<span style=\"{style}\">{encoded}</span>";
    }

    private static string BuildRunStyle(A.RunProperties? runProperties, A.DefaultRunProperties? defaultRunProperties, A.ColorScheme? colorScheme)
    {
        var fontSize = GetInt(runProperties?.FontSize) ?? GetInt(defaultRunProperties?.FontSize);
        var fontSizePx = EmuFontToPx(fontSize);
        var fontFamily = GetFontFamily(runProperties, defaultRunProperties);
        var color = ResolveColor(runProperties, defaultRunProperties, colorScheme);
        var bold = GetBool(runProperties?.Bold) ?? GetBool(defaultRunProperties?.Bold) ?? false;
        var italic = GetBool(runProperties?.Italic) ?? GetBool(defaultRunProperties?.Italic) ?? false;
        var underlineValue = NormalizeToken(GetStringValue(runProperties?.Underline) ?? GetStringValue(defaultRunProperties?.Underline));

        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(fontFamily))
        {
            parts.Add($"font-family:'{fontFamily}', 'Segoe UI', sans-serif");
        }
        if (fontSizePx > 0)
        {
            parts.Add($"font-size:calc(var(--scale,1) * {fontSizePx:0.###}px)");
        }
        if (!string.IsNullOrWhiteSpace(color))
        {
            parts.Add($"color:{color}");
        }
        if (bold)
        {
            parts.Add("font-weight:700");
        }
        if (italic)
        {
            parts.Add("font-style:italic");
        }
        if (!string.IsNullOrWhiteSpace(underlineValue) &&
            underlineValue != "none" &&
            underlineValue != "false" &&
            underlineValue != "0")
        {
            parts.Add("text-decoration:underline");
        }

        return string.Join(";", parts);
    }

    private static string? GetFontFamily(A.RunProperties? runProperties, A.DefaultRunProperties? defaultRunProperties)
    {
        var family = GetTypeface(runProperties) ?? GetTypeface(defaultRunProperties);

        if (string.IsNullOrWhiteSpace(family) || family.StartsWith("+", StringComparison.Ordinal))
        {
            return null;
        }

        return family;
    }

    private static string? GetTypeface(OpenXmlElement? runProperties)
    {
        if (runProperties is null)
        {
            return null;
        }

        var latin = runProperties.GetFirstChild<A.LatinFont>();
        var eastAsian = runProperties.GetFirstChild<A.EastAsianFont>();

        return GetStringValue(latin?.Typeface) ?? GetStringValue(eastAsian?.Typeface);
    }

    private static string ResolveColor(A.RunProperties? runProperties, A.DefaultRunProperties? defaultRunProperties, A.ColorScheme? colorScheme)
    {
        var fill = GetSolidFill(runProperties) ?? GetSolidFill(defaultRunProperties);
        if (fill is null)
        {
            return "#000000";
        }

        if (fill.RgbColorModelHex is not null)
        {
            var hex = GetStringValue(fill.RgbColorModelHex.Val);
            if (!string.IsNullOrWhiteSpace(hex))
            {
                return $"#{hex}";
            }
        }

        if (fill.SchemeColor is not null && colorScheme is not null)
        {
            var hex = ResolveSchemeColor(fill.SchemeColor.Val, colorScheme);
            if (!string.IsNullOrWhiteSpace(hex))
            {
                return $"#{hex}";
            }
        }

        return "#000000";
    }

    private static A.SolidFill? GetSolidFill(OpenXmlElement? runProperties)
    {
        return runProperties?.GetFirstChild<A.SolidFill>();
    }

    private static int? GetInt(OpenXmlSimpleType? value)
    {
        var text = GetStringValue(value);
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : null;
    }

    private static long? GetLong(OpenXmlSimpleType? value)
    {
        var text = GetStringValue(value);
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        return long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : null;
    }

    private static bool? GetBool(OpenXmlSimpleType? value)
    {
        var text = GetStringValue(value);
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        if (text == "1")
        {
            return true;
        }
        if (text == "0")
        {
            return false;
        }

        return bool.TryParse(text, out var parsed) ? parsed : null;
    }

    private static string? GetStringValue(OpenXmlSimpleType? value)
    {
        if (value is null)
        {
            return null;
        }

        var text = value.InnerText;
        return string.IsNullOrWhiteSpace(text) ? null : text;
    }

    private static string? NormalizeToken(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return value.Trim()
            .Replace(" ", string.Empty, StringComparison.Ordinal)
            .Replace("_", string.Empty, StringComparison.Ordinal)
            .ToLowerInvariant();
    }

    private static string? ResolveSchemeColor(OpenXmlSimpleType? schemeValue, A.ColorScheme colorScheme)
    {
        var token = NormalizeToken(GetStringValue(schemeValue));
        if (string.IsNullOrWhiteSpace(token))
        {
            return null;
        }

        OpenXmlElement? color = token switch
        {
            "dk1" or "dark1" => colorScheme.Dark1Color,
            "lt1" or "light1" => colorScheme.Light1Color,
            "dk2" or "dark2" => colorScheme.Dark2Color,
            "lt2" or "light2" => colorScheme.Light2Color,
            "accent1" => colorScheme.Accent1Color,
            "accent2" => colorScheme.Accent2Color,
            "accent3" => colorScheme.Accent3Color,
            "accent4" => colorScheme.Accent4Color,
            "accent5" => colorScheme.Accent5Color,
            "accent6" => colorScheme.Accent6Color,
            "hlink" or "hyperlink" => colorScheme.Hyperlink,
            "folhlink" or "followedhyperlink" => colorScheme.FollowedHyperlinkColor,
            _ => null
        };

        if (color is null)
        {
            return null;
        }

        var rgb = color.Descendants<A.RgbColorModelHex>().FirstOrDefault();
        return GetStringValue(rgb?.Val);
    }

    private static double EmuFontToPx(int? fontSize)
    {
        if (fontSize is null || fontSize <= 0)
        {
            return 16;
        }

        var points = fontSize.Value / 100.0;
        return points * 96.0 / 72.0;
    }

    private static double EmuToPx(long emu)
    {
        const double emuPerInch = 914400.0;
        return emu / emuPerInch * 96.0;
    }

    private static string GetImageExtension(ImagePart imagePart)
    {
        var extension = Path.GetExtension(imagePart.Uri.OriginalString);
        if (!string.IsNullOrWhiteSpace(extension))
        {
            return extension;
        }

        return imagePart.ContentType switch
        {
            "image/png" => ".png",
            "image/jpeg" => ".jpg",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            "image/tiff" => ".tif",
            "image/x-emf" => ".emf",
            "image/x-wmf" => ".wmf",
            _ => ".bin"
        };
    }

    private static string GetPartExtension(OpenXmlPart part)
    {
        var extension = Path.GetExtension(part.Uri.OriginalString);
        if (!string.IsNullOrWhiteSpace(extension))
        {
            return extension;
        }

        return part.ContentType switch
        {
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document" => ".docx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" => ".xlsx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation" => ".pptx",
            "application/pdf" => ".pdf",
            "application/zip" => ".zip",
            "application/octet-stream" => ".bin",
            "audio/mpeg" => ".mp3",
            "audio/wav" => ".wav",
            "audio/x-wav" => ".wav",
            "audio/mp4" => ".m4a",
            "video/mp4" => ".mp4",
            "video/quicktime" => ".mov",
            "video/x-msvideo" => ".avi",
            _ => ".bin"
        };
    }

    private static string ComputeSha256Hex(byte[] data)
    {
        using var sha = SHA256.Create();
        var hash = sha.ComputeHash(data);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    private static string BuildHtmlFromLibreOffice(
        string html,
        string title,
        string outputDir,
        SlideSizeInfo slideSizeInfo,
        List<HtmlSlide> slides,
        List<HtmlImageBlock> logos,
        List<HtmlAttachment> attachments)
    {
        var bodyMatch = Regex.Match(html, "<body[^>]*>", RegexOptions.IgnoreCase);
        var bodyCloseMatch = Regex.Match(html, "</body>", RegexOptions.IgnoreCase);
        if (!bodyMatch.Success || !bodyCloseMatch.Success || bodyCloseMatch.Index <= bodyMatch.Index)
        {
            return html;
        }

        var slideWidthPx = EmuToPx(slideSizeInfo.WidthEmu);
        var slideHeightPx = EmuToPx(slideSizeInfo.HeightEmu);
        if (slideWidthPx <= 0 || slideHeightPx <= 0)
        {
            slideWidthPx = 960;
            slideHeightPx = 540;
        }

        var bodyStart = bodyMatch.Index + bodyMatch.Length;
        var bodyInner = html.Substring(bodyStart, bodyCloseMatch.Index - bodyStart);

        bodyInner = RemoveEmptyParagraphs(bodyInner);

        bodyInner = Regex.Replace(
            bodyInner,
            "<h1[^>]*page-break-before[^>]*>(?<content>.*?)</h1>",
            match =>
            {
                var content = match.Groups["content"].Value;
                if (string.IsNullOrWhiteSpace(content))
                {
                    return "<!--SLIDE_BREAK-->";
                }

                return $"<!--SLIDE_BREAK--><h1>{content}</h1>";
            },
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

        var segments = bodyInner.Split("<!--SLIDE_BREAK-->", StringSplitOptions.None).ToList();
        var slideCount = slides.Count;
        if (segments.Count < slideCount)
        {
            while (segments.Count < slideCount)
            {
                segments.Add(string.Empty);
            }
        }
        else if (segments.Count > slideCount && slideCount > 0)
        {
            var overflow = string.Join(string.Empty, segments.Skip(slideCount));
            segments = segments.Take(slideCount).ToList();
            segments[^1] += overflow;
        }

        var safeTitle = WebUtility.HtmlEncode(title);
        var sb = new StringBuilder();
        sb.Append("<main class=\"document\">");

        if (logos.Count > 0)
        {
            sb.Append("<header class=\"logos\">");
            foreach (var logo in logos)
            {
                var relativePath = Path.GetRelativePath(outputDir, logo.FilePath);
                var htmlPath = relativePath.Replace('\\', '/');
                sb.Append($"<img src=\"{htmlPath}\" alt=\"{safeTitle} logo\" loading=\"lazy\">");
            }
            sb.Append("</header>");
        }

        for (var i = 0; i < slideCount; i++)
        {
            var slide = slides[i];
            var background = string.IsNullOrWhiteSpace(slide.BackgroundColor) ? "#ffffff" : slide.BackgroundColor;
            var (backgroundImages, flowImages) = SplitImageLayers(slide.Images, slideSizeInfo);
            var contentHtml = InsertImagesIntoSegment(segments[i], flowImages, slideSizeInfo, outputDir, safeTitle);
            var backgroundHtml = BuildBackgroundImagesHtml(backgroundImages, outputDir, safeTitle);
            var minHeightPx = GetMaxBottomPx(backgroundImages);
            var minHeightStyle = minHeightPx > 0
                ? $"min-height:calc(var(--scale,1) * {FormatNumber(minHeightPx)}px);"
                : string.Empty;

            sb.Append($"<section class=\"slide\" data-slide=\"{slide.Index:000}\" data-base-width=\"{FormatNumber(slideWidthPx)}\" data-base-height=\"{FormatNumber(slideHeightPx)}\" style=\"background:{background};{minHeightStyle}\">");
            if (!string.IsNullOrWhiteSpace(backgroundHtml))
            {
                sb.Append("<div class=\"slide-background\">");
                sb.Append(backgroundHtml);
                sb.Append("</div>");
            }
            sb.Append("<div class=\"slide-content\">");
            sb.Append(contentHtml);
            sb.Append("</div>");
            sb.Append("</section>");
        }

        if (attachments.Count > 0)
        {
            sb.Append(BuildAttachmentsHtml(attachments, outputDir));
        }

        sb.Append("</main>");
        sb.Append("<script>");
        sb.Append("const updateScale=()=>{document.querySelectorAll('.slide').forEach(slide=>{const baseWidth=parseFloat(slide.dataset.baseWidth||'0');if(!baseWidth)return;const scale=slide.clientWidth/baseWidth;slide.style.setProperty('--scale',scale.toFixed(4));});};");
        sb.Append("window.addEventListener('resize',updateScale);window.addEventListener('load',updateScale);updateScale();");
        sb.Append("</script>");

        var rebuilt = html.Substring(0, bodyStart) + sb + html.Substring(bodyCloseMatch.Index);

        var styleBlock = new StringBuilder();
        styleBlock.AppendLine("<style>");
        styleBlock.AppendLine("  body { margin: 0; padding: 0; font-family: \"Segoe UI\", Arial, sans-serif; background: #ffffff; }");
        styleBlock.AppendLine("  .document { max-width: 1200px; margin: 0 auto; padding: 24px; display: block; }");
        styleBlock.AppendLine("  .slide { position: relative; width: 100%; overflow: hidden; margin-bottom: 32px; }");
        styleBlock.AppendLine("  .slide-background { position: absolute; inset: 0; z-index: 1; pointer-events: none; }");
        styleBlock.AppendLine("  .slide-content { position: relative; z-index: 2; }");
        styleBlock.AppendLine("  .pptx-flow-image { display: block; max-width: 100%; height: auto; margin: 0; }");
        styleBlock.AppendLine("  .pptx-abs-image { position: absolute; max-width: 100%; height: 100%; object-fit: contain; }");
        styleBlock.AppendLine("  img { max-width: 100%; height: auto; }");
        styleBlock.AppendLine("  table { width: 100%; border-collapse: collapse; }");
        styleBlock.AppendLine("  td, th { vertical-align: top; }");
        styleBlock.AppendLine("  .logos { display: flex; flex-wrap: wrap; align-items: center; justify-content: center; gap: 16px; padding: 16px; margin-bottom: 24px; }");
        styleBlock.AppendLine("  .logos img { max-height: 120px; max-width: 40vw; height: auto; width: auto; object-fit: contain; }");
        styleBlock.AppendLine("  .attachments { padding: 16px 20px; }");
        styleBlock.AppendLine("  .attachments h2 { margin: 0 0 12px; font-size: 18px; }");
        styleBlock.AppendLine("  .attachments ul { margin: 0; padding-left: 20px; }");
        styleBlock.AppendLine("</style>");

        return InsertStyleBlock(rebuilt, styleBlock.ToString());
    }

    private static string InsertStyleBlock(string html, string styleBlock)
    {
        var headCloseMatch = Regex.Match(html, "</head>", RegexOptions.IgnoreCase);
        if (headCloseMatch.Success)
        {
            return html.Insert(headCloseMatch.Index, styleBlock);
        }

        return styleBlock + html;
    }

    private static string RemoveEmptyParagraphs(string html)
    {
        if (string.IsNullOrWhiteSpace(html))
        {
            return html;
        }

        var cleaned = Regex.Replace(
            html,
            "<p>\\s*(?:&nbsp;|&#160;|<br\\s*/?>|\\s)*</p>",
            string.Empty,
            RegexOptions.IgnoreCase);

        return cleaned;
    }

    private static string InsertImagesIntoSegment(
        string segmentHtml,
        List<HtmlImageBlock> images,
        SlideSizeInfo slideSizeInfo,
        string outputDir,
        string safeTitle)
    {
        if (images.Count == 0)
        {
            return segmentHtml;
        }

        var blocks = SplitBlocks(segmentHtml);
        var slideHeight = Math.Max(1, slideSizeInfo.HeightEmu);

        foreach (var image in images.OrderBy(img => img.Y).ThenBy(img => img.X))
        {
            var positionRatio = (image.Y + (image.Cy / 2.0)) / slideHeight;
            var index = blocks.Count == 0
                ? 0
                : (int)Math.Round(positionRatio * blocks.Count, MidpointRounding.AwayFromZero);
            index = Math.Clamp(index, 0, blocks.Count);

            var imageHtml = BuildFlowImageTag(image, outputDir, safeTitle);
            blocks.Insert(index, imageHtml);
        }

        return string.Concat(blocks);
    }

    private static (List<HtmlImageBlock> Background, List<HtmlImageBlock> Flow) SplitImageLayers(
        List<HtmlImageBlock> images,
        SlideSizeInfo slideSizeInfo)
    {
        if (images.Count == 0)
        {
            return (new List<HtmlImageBlock>(), new List<HtmlImageBlock>());
        }

        var background = new List<HtmlImageBlock>();
        var flow = new List<HtmlImageBlock>();

        var slideWidth = Math.Max(1, slideSizeInfo.WidthEmu);
        var slideHeight = Math.Max(1, slideSizeInfo.HeightEmu);
        var slideArea = (double)slideWidth * slideHeight;

        foreach (var image in images)
        {
            var widthRatio = image.Cx / (double)slideWidth;
            var heightRatio = image.Cy / (double)slideHeight;
            var areaRatio = slideArea <= 0 ? 0 : (image.Cx * (double)image.Cy) / slideArea;

            var isBackgroundCandidate = widthRatio >= 0.9 ||
                heightRatio >= 0.9 ||
                areaRatio >= 0.7;

            if (isBackgroundCandidate)
            {
                background.Add(image);
            }
            else
            {
                flow.Add(image);
            }
        }

        return (background, flow);
    }

    private static string BuildBackgroundImagesHtml(
        List<HtmlImageBlock> images,
        string outputDir,
        string safeTitle)
    {
        if (images.Count == 0)
        {
            return string.Empty;
        }

        var sb = new StringBuilder();
        foreach (var image in images.OrderBy(img => img.Y).ThenBy(img => img.X))
        {
            sb.Append(BuildAbsoluteImageTag(image, outputDir, safeTitle));
        }

        return sb.ToString();
    }

    private static double GetMaxBottomPx(List<HtmlImageBlock> images)
    {
        if (images.Count == 0)
        {
            return 0;
        }

        var maxBottomEmu = images.Max(image => image.Y + image.Cy);
        return EmuToPx(maxBottomEmu);
    }

    private static List<string> SplitBlocks(string html)
    {
        var blocks = new List<string>();
        if (string.IsNullOrWhiteSpace(html))
        {
            return blocks;
        }

        var regex = new Regex("<(/?)(h[1-6]|p|ul|ol|table)[^>]*>", RegexOptions.IgnoreCase);
        var lastIndex = 0;
        var current = new StringBuilder();
        var stack = new Stack<string>();

        foreach (Match match in regex.Matches(html))
        {
            var end = match.Index + match.Length;
            current.Append(html.Substring(lastIndex, end - lastIndex));
            lastIndex = end;

            var isClosing = match.Groups[1].Value == "/";
            var tag = match.Groups[2].Value.ToLowerInvariant();

            if (!isClosing)
            {
                if (tag is "ul" or "ol" or "table")
                {
                    stack.Push(tag);
                }

                continue;
            }

            if (tag is "ul" or "ol" or "table")
            {
                if (stack.Count > 0)
                {
                    stack.Pop();
                }

                if (stack.Count == 0)
                {
                    blocks.Add(current.ToString());
                    current.Clear();
                }

                continue;
            }

            if (tag == "p" || tag.StartsWith("h", StringComparison.Ordinal))
            {
                if (stack.Count == 0)
                {
                    blocks.Add(current.ToString());
                    current.Clear();
                }
            }
        }

        if (lastIndex < html.Length)
        {
            current.Append(html.Substring(lastIndex));
        }

        if (current.Length > 0)
        {
            blocks.Add(current.ToString());
        }

        return blocks;
    }

    private static string BuildFlowImageTag(
        HtmlImageBlock image,
        string outputDir,
        string safeTitle)
    {
        var relativePath = Path.GetRelativePath(outputDir, image.FilePath);
        var htmlPath = relativePath.Replace('\\', '/');
        var xPx = EmuToPx(image.X);
        var widthPx = EmuToPx(image.Cx);

        var style = $"margin-left:calc(var(--scale,1) * {FormatNumber(xPx)}px);" +
                    $"width:calc(var(--scale,1) * {FormatNumber(widthPx)}px);" +
                    "max-width:100%;height:auto;display:block;";

        return $"<img class=\"pptx-flow-image\" src=\"{htmlPath}\" alt=\"{safeTitle}\" loading=\"lazy\" style=\"{style}\">";
    }

    private static string BuildAbsoluteImageTag(
        HtmlImageBlock image,
        string outputDir,
        string safeTitle)
    {
        var relativePath = Path.GetRelativePath(outputDir, image.FilePath);
        var htmlPath = relativePath.Replace('\\', '/');
        var xPx = EmuToPx(image.X);
        var yPx = EmuToPx(image.Y);
        var widthPx = EmuToPx(image.Cx);
        var heightPx = EmuToPx(image.Cy);

        var style = $"left:calc(var(--scale,1) * {FormatNumber(xPx)}px);" +
                    $"top:calc(var(--scale,1) * {FormatNumber(yPx)}px);" +
                    $"width:calc(var(--scale,1) * {FormatNumber(widthPx)}px);" +
                    $"height:calc(var(--scale,1) * {FormatNumber(heightPx)}px);";

        return $"<img class=\"pptx-abs-image\" src=\"{htmlPath}\" alt=\"{safeTitle}\" loading=\"lazy\" style=\"{style}\">";
    }

    private static string BuildAttachmentsHtml(List<HtmlAttachment> attachments, string outputDir)
    {
        if (attachments.Count == 0)
        {
            return string.Empty;
        }

        var sb = new StringBuilder();
        sb.Append("<section class=\"attachments\">");
        sb.Append("<h2>Přílohy</h2>");
        sb.Append("<ul>");
        foreach (var attachment in attachments.OrderBy(item => item.SlideIndex).ThenBy(item => item.FilePath, StringComparer.OrdinalIgnoreCase))
        {
            var relativePath = Path.GetRelativePath(outputDir, attachment.FilePath);
            var htmlPath = relativePath.Replace('\\', '/');
            var name = WebUtility.HtmlEncode(Path.GetFileName(attachment.FilePath));
            sb.Append($"<li><a href=\"{htmlPath}\" target=\"_blank\" rel=\"noopener\">{name}</a> (slide {attachment.SlideIndex:000})</li>");
        }
        sb.Append("</ul>");
        sb.Append("</section>");
        return sb.ToString();
    }

    private static string BuildPositionStyle(long x, long y, long cx, long cy, SlideSizeInfo slideSizeInfo)
    {
        var widthEmu = Math.Max(1, slideSizeInfo.WidthEmu);
        var heightEmu = Math.Max(1, slideSizeInfo.HeightEmu);

        var left = x / (double)widthEmu * 100.0;
        var top = y / (double)heightEmu * 100.0;
        var width = cx / (double)widthEmu * 100.0;
        var height = cy / (double)heightEmu * 100.0;

        return $"left:{FormatNumber(left)}%;top:{FormatNumber(top)}%;width:{FormatNumber(width)}%;height:{FormatNumber(height)}%;";
    }

    private static string FormatNumber(double value)
    {
        return value.ToString("0.####", CultureInfo.InvariantCulture);
    }

    private static string BuildPaddingStyle(HtmlPadding padding)
    {
        if (padding.LeftPx <= 0 && padding.TopPx <= 0 && padding.RightPx <= 0 && padding.BottomPx <= 0)
        {
            return string.Empty;
        }

        return $"padding:calc(var(--scale,1) * {padding.TopPx:0.###}px) calc(var(--scale,1) * {padding.RightPx:0.###}px) calc(var(--scale,1) * {padding.BottomPx:0.###}px) calc(var(--scale,1) * {padding.LeftPx:0.###}px);";
    }

    private static bool IsLogoCandidate(HtmlImageBlock image, SlideSizeInfo slideSizeInfo)
    {
        if (slideSizeInfo.WidthEmu <= 0 || slideSizeInfo.HeightEmu <= 0)
        {
            return true;
        }

        var widthRatio = image.Cx / (double)slideSizeInfo.WidthEmu;
        var heightRatio = image.Cy / (double)slideSizeInfo.HeightEmu;
        if (widthRatio > 0.25 || heightRatio > 0.25)
        {
            return false;
        }

        var edgeThresholdX = slideSizeInfo.WidthEmu * 0.10;
        var edgeThresholdY = slideSizeInfo.HeightEmu * 0.10;

        var leftEdge = image.X <= edgeThresholdX;
        var rightEdge = image.X + image.Cx >= slideSizeInfo.WidthEmu - edgeThresholdX;
        var topEdge = image.Y <= edgeThresholdY;
        var bottomEdge = image.Y + image.Cy >= slideSizeInfo.HeightEmu - edgeThresholdY;

        return (topEdge || bottomEdge) && (leftEdge || rightEdge);
    }
}
