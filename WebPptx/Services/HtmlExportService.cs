using System.Net;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Extensions.Hosting;
using WebPptx.Controllers;

namespace WebPptx.Services;

public interface IHtmlExportService
{
    ExportHtmlResponse Export(string inputPath, ExportHtmlRequest request);
}

public class HtmlExportService : IHtmlExportService
{
    private readonly IConfiguration _configuration;
    private readonly IHostEnvironment _environment;

    public HtmlExportService(
        IConfiguration configuration,
        IHostEnvironment environment)
    {
        _configuration = configuration;
        _environment = environment;
    }

    public ExportHtmlResponse Export(string inputPath, ExportHtmlRequest request)
    {
        var options = BuildHtmlExportOptions(request);
        return ExportHtmlSingle(inputPath, options);
    }

    private HtmlExportOptions BuildHtmlExportOptions(ExportHtmlRequest request)
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

        var frameMaxWidth = request.FrameMaxWidth
            ?? _configuration.GetValue<int?>("Screenshots:FrameMaxWidth")
            ?? maxWidth;
        var frameMaxHeight = request.FrameMaxHeight
            ?? _configuration.GetValue<int?>("Screenshots:FrameMaxHeight")
            ?? maxHeight;
        var frameAllowUpscale = request.FrameAllowUpscale
            ?? _configuration.GetValue("Screenshots:FrameAllowUpscale", true);

        var configuredSofficePath = string.IsNullOrWhiteSpace(request.SofficePath)
            ? _configuration["LibreOffice:SofficePath"]
            : request.SofficePath;

        return new HtmlExportOptions(
            configuredSofficePath,
            maxWidth,
            maxHeight,
            jpegQuality,
            pdfDpi,
            frameMaxWidth,
            frameMaxHeight,
            frameAllowUpscale);
    }

    private ExportHtmlResponse ExportHtmlSingle(string inputPath, HtmlExportOptions options)
    {
        if (string.IsNullOrWhiteSpace(inputPath))
        {
            throw new ArgumentException("Path must be a non-empty string.");
        }

        var fullPath = Path.GetFullPath(inputPath);
        if (!System.IO.File.Exists(fullPath))
        {
            throw new FileNotFoundException("File not found.", fullPath);
        }

        if (!string.Equals(Path.GetExtension(fullPath), ".pptx", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Only .pptx files are supported.");
        }

        var samplesRoot = GetSamplesRoot();
        Directory.CreateDirectory(samplesRoot);

        var outputDir = Path.Combine(samplesRoot, Path.GetFileNameWithoutExtension(fullPath));
        Directory.CreateDirectory(outputDir);

        var framesDir = Path.Combine(outputDir, "frames");
        EnsureEmptyDirectory(framesDir);

        var slideCount = 0;
        SlideSizeInfo? slideSizeInfo = null;
        Dictionary<int, List<SlideFrameRect>>? framesBySlide = null;

        using (var presentation = PresentationDocument.Open(fullPath, false))
        {
            var presentationPart = presentation.PresentationPart;
            var slideIds = presentationPart?.Presentation?.SlideIdList?.ChildElements.OfType<SlideId>().ToList();

            if (presentationPart is not null && slideIds is not null)
            {
                slideCount = slideIds.Count;
                slideSizeInfo = PptxController.GetSlideSizeInfo(presentationPart);
                framesBySlide = PptxController.ExtractSlideFrames(presentationPart, slideIds);
            }
        }

        var sofficePath = PptxController.ResolveSofficePath(options.SofficePath);
        if (sofficePath is null)
        {
            throw new InvalidOperationException("LibreOffice (soffice) not found. Provide 'sofficePath' or set LibreOffice:SofficePath.");
        }

        var frameMetadata = new List<FrameScreenshotInfo>();
        var screenshotFiles = PptxController.ExtractSlideScreenshotsViaPdf(
            fullPath,
            framesDir,
            sofficePath,
            slideCount,
            options.MaxWidth,
            options.MaxHeight,
            options.JpegQuality,
            options.PdfDpi,
            true,
            framesBySlide,
            slideSizeInfo,
            options.FrameMaxWidth,
            options.FrameMaxHeight,
            options.FrameAllowUpscale,
            frameMetadata);

        var orderedFrames = GetOrderedHtmlFrames(frameMetadata, screenshotFiles);
        var htmlPath = Path.Combine(outputDir, "index.html");
        var title = Path.GetFileNameWithoutExtension(fullPath);
        var html = BuildHtmlDocument(title, outputDir, orderedFrames);
        System.IO.File.WriteAllText(htmlPath, html);

        var frameFiles = orderedFrames.Select(frame => frame.FilePath).ToList();
        return new ExportHtmlResponse(fullPath, outputDir, htmlPath, slideCount, frameFiles.Count, frameFiles);
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

    private static List<FrameScreenshotInfo> GetOrderedHtmlFrames(
        List<FrameScreenshotInfo> frameMetadata,
        List<string> screenshotFiles)
    {
        var frames = frameMetadata
            .Where(frame => string.Equals(frame.Type, "frame", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(frame.Type, "slide", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (frames.Count == 0 && screenshotFiles.Count > 0)
        {
            frames = BuildFallbackFrameMetadata(screenshotFiles);
        }

        return frames
            .OrderBy(frame => frame.SlideIndex)
            .ThenBy(frame => frame.Y)
            .ThenBy(frame => frame.X)
            .ThenBy(frame => frame.FrameIndex)
            .ToList();
    }

    private static List<FrameScreenshotInfo> BuildFallbackFrameMetadata(IEnumerable<string> files)
    {
        var frames = new List<FrameScreenshotInfo>();

        foreach (var file in files)
        {
            var name = Path.GetFileNameWithoutExtension(file);
            if (!name.StartsWith("slide-", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (name.Length < 10)
            {
                continue;
            }

            var slideToken = name.Substring(6, 3);
            if (!int.TryParse(slideToken, out var slideIndex))
            {
                continue;
            }

            if (name.Contains("-frame", StringComparison.OrdinalIgnoreCase))
            {
                var frameTokenIndex = name.IndexOf("-frame", StringComparison.OrdinalIgnoreCase) + 6;
                var frameToken = name[frameTokenIndex..];
                if (!int.TryParse(frameToken, out var frameIndex))
                {
                    frameIndex = 0;
                }

                frames.Add(new FrameScreenshotInfo(slideIndex, frameIndex, file, 0, 0, 0, 0, "frame"));
            }
            else if (name.EndsWith("-screenshot", StringComparison.OrdinalIgnoreCase))
            {
                frames.Add(new FrameScreenshotInfo(slideIndex, 0, file, 0, 0, 0, 0, "slide"));
            }
        }

        return frames;
    }

    private static string BuildHtmlDocument(
        string title,
        string outputDir,
        List<FrameScreenshotInfo> orderedFrames)
    {
        var safeTitle = WebUtility.HtmlEncode(title);
        var sb = new StringBuilder();
        sb.AppendLine("<!doctype html>");
        sb.AppendLine("<html lang=\"cs\">");
        sb.AppendLine("<head>");
        sb.AppendLine("  <meta charset=\"utf-8\">");
        sb.AppendLine("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">");
        sb.AppendLine($"  <title>{safeTitle}</title>");
        sb.AppendLine("  <style>");
        sb.AppendLine("    :root {");
        sb.AppendLine("      --bg: #f5f3ee;");
        sb.AppendLine("      --card: #ffffff;");
        sb.AppendLine("      --border: #dedad4;");
        sb.AppendLine("      --text: #1f1f1f;");
        sb.AppendLine("      --muted: #6b6761;");
        sb.AppendLine("      --gap: 16px;");
        sb.AppendLine("    }");
        sb.AppendLine("    * { box-sizing: border-box; }");
        sb.AppendLine("    body {");
        sb.AppendLine("      margin: 0;");
        sb.AppendLine("      font-family: \"Fira Sans\", \"Segoe UI\", sans-serif;");
        sb.AppendLine("      background: var(--bg);");
        sb.AppendLine("      color: var(--text);");
        sb.AppendLine("    }");
        sb.AppendLine("    header { padding: 24px 24px 8px; }");
        sb.AppendLine("    header h1 { margin: 0 0 6px; font-size: 28px; }");
        sb.AppendLine("    header p { margin: 0; color: var(--muted); }");
        sb.AppendLine("    main { padding: 0 24px 32px; }");
        sb.AppendLine("    .doc-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: var(--gap); }");
        sb.AppendLine("    .slide-break { grid-column: 1 / -1; margin-top: 12px; padding: 10px 12px; background: var(--card); border: 1px solid var(--border); border-radius: 10px; color: var(--muted); font-size: 14px; }");
        sb.AppendLine("    .frame { background: #fff; border: 1px solid var(--border); border-radius: 10px; overflow: hidden; }");
        sb.AppendLine("    .frame img { width: 100%; height: auto; display: block; }");
        sb.AppendLine("    .empty { padding: 24px; background: var(--card); border: 1px dashed var(--border); border-radius: 12px; color: var(--muted); }");
        sb.AppendLine("  </style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine($"  <header><h1>{safeTitle}</h1><p>Generated from PPTX frames</p></header>");
        sb.AppendLine("  <main>");

        if (orderedFrames.Count == 0)
        {
            sb.AppendLine("    <div class=\"empty\">No frames were exported.</div>");
        }
        else
        {
            sb.AppendLine("    <div class=\"doc-grid\">");
            int? currentSlide = null;
            foreach (var frame in orderedFrames)
            {
                if (currentSlide != frame.SlideIndex)
                {
                    currentSlide = frame.SlideIndex;
                    sb.AppendLine($"      <div class=\"slide-break\">Slide {currentSlide:000}</div>");
                }

                var relativePath = Path.GetRelativePath(outputDir, frame.FilePath);
                var htmlPath = relativePath.Replace('\\', '/');
                var alt = $"Slide {frame.SlideIndex:000} frame {frame.FrameIndex:00}";
                sb.AppendLine($"      <div class=\"frame\" data-slide=\"{frame.SlideIndex:000}\" data-frame=\"{frame.FrameIndex:00}\">");
                sb.AppendLine($"        <img src=\"{htmlPath}\" alt=\"{alt}\" loading=\"lazy\">");
                sb.AppendLine("      </div>");
            }
            sb.AppendLine("    </div>");
        }

        sb.AppendLine("  </main>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");

        return sb.ToString();
    }
}
