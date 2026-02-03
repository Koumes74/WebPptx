using System.Diagnostics;
using System.IO.Compression;
using System.Text.Json;
using WebPptx.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.AspNetCore.Mvc;
using PdfiumViewer;
using SixLabors.ImageSharp;
using ImageSharp = SixLabors.ImageSharp.Image;
using SixLabors.ImageSharp.Formats.Jpeg;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using DrawingImageFormat = System.Drawing.Imaging.ImageFormat;
using ImageSharpRectangle = SixLabors.ImageSharp.Rectangle;
using A = DocumentFormat.OpenXml.Drawing;

namespace WebPptx.Controllers;

[ApiController]
[Route("pptx")]
public class PptxController : ControllerBase
{
    private readonly ILogger<PptxController> _logger;
    private readonly IConfiguration _configuration;

    public PptxController(ILogger<PptxController> logger, IConfiguration configuration)
    {
        _logger = logger;
        _configuration = configuration;
    }

    [HttpPost("extract")]
    [ProducesResponseType(typeof(ExtractPptxResponse), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(ExtractPptxBatchResponse), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status500InternalServerError)]
    public IActionResult Extract([FromBody] ExtractPptxRequest request)
    {
        if (request is null)
        {
            return BadRequest("Body is required.");
        }

        List<string> requestedPaths;
        try
        {
            requestedPaths = GetRequestedPaths(request);
        }
        catch (ArgumentException ex)
        {
            return BadRequest(ex.Message);
        }
        if (requestedPaths.Count == 0)
        {
            if (HasDirectories(request))
            {
                return BadRequest("No .pptx files found in the provided directory.");
            }

            return BadRequest("Body must contain a non-empty 'path', 'paths', 'directory', or 'directories' value.");
        }

        var options = BuildExtractOptions(request);

        if (requestedPaths.Count == 1)
        {
            try
            {
                var singleResult = ExtractSingle(requestedPaths[0], options);
                return Ok(singleResult);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (FileNotFoundException ex)
            {
                return NotFound($"File not found: {ex.FileName ?? ex.Message}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Extraction failed for {Path}", requestedPaths[0]);
                return Problem(title: "Extraction failed", detail: ex.Message, statusCode: StatusCodes.Status500InternalServerError);
            }
        }

        var items = new List<ExtractPptxItemResult>();
        foreach (var path in requestedPaths)
        {
            try
            {
                var result = ExtractSingle(path, options);
                items.Add(new ExtractPptxItemResult(path, true, null, result));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Extraction failed for {Path}", path);
                items.Add(new ExtractPptxItemResult(path, false, ex.Message, null));
            }
        }

        var succeeded = items.Count(item => item.Success);
        var failed = items.Count - succeeded;
        return Ok(new ExtractPptxBatchResponse(items.Count, succeeded, failed, items));
    }

    [HttpPost("rebuild")]
    [ProducesResponseType(typeof(RebuildPptxResponse), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    [ProducesResponseType(StatusCodes.Status500InternalServerError)]
    public IActionResult Rebuild([FromBody] RebuildPptxRequest request, [FromServices] IPptxRebuildService rebuildService)
    {
        try
        {
            var result = rebuildService.RebuildFromFrames(request);
            return Ok(new RebuildPptxResponse(result.OutputPath, result.SlideCount, result.ItemCount));
        }
        catch (ArgumentException ex)
        {
            return BadRequest(ex.Message);
        }
        catch (FileNotFoundException ex)
        {
            return BadRequest(ex.Message);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Rebuild failed.");
            return Problem(title: "Rebuild failed", detail: ex.Message, statusCode: StatusCodes.Status500InternalServerError);
        }
    }

    [HttpPost("export-html")]
    [ProducesResponseType(typeof(ExportHtmlResponse), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(ExportHtmlBatchResponse), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status500InternalServerError)]
    public IActionResult ExportHtml([FromBody] ExportHtmlRequest request, [FromServices] IHtmlExportService htmlExportService)
    {
        if (request is null)
        {
            return BadRequest("Body is required.");
        }

        List<string> requestedPaths;
        try
        {
            requestedPaths = GetRequestedHtmlPaths(request);
        }
        catch (ArgumentException ex)
        {
            return BadRequest(ex.Message);
        }

        if (requestedPaths.Count == 0)
        {
            return BadRequest("Body must contain a non-empty 'path' or 'paths' value.");
        }

        if (requestedPaths.Count == 1)
        {
            try
            {
                var result = htmlExportService.Export(requestedPaths[0], request);
                return Ok(result);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (FileNotFoundException ex)
            {
                return NotFound($"File not found: {ex.FileName ?? ex.Message}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "HTML export failed for {Path}", requestedPaths[0]);
                return Problem(title: "HTML export failed", detail: ex.Message, statusCode: StatusCodes.Status500InternalServerError);
            }
        }

        var items = new List<ExportHtmlItemResult>();
        foreach (var path in requestedPaths)
        {
            try
            {
                var result = htmlExportService.Export(path, request);
                items.Add(new ExportHtmlItemResult(path, true, null, result));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "HTML export failed for {Path}", path);
                items.Add(new ExportHtmlItemResult(path, false, ex.Message, null));
            }
        }

        var succeeded = items.Count(item => item.Success);
        var failed = items.Count - succeeded;
        return Ok(new ExportHtmlBatchResponse(items.Count, succeeded, failed, items));
    }

    [HttpPost("htmlpage")]
    [ProducesResponseType(typeof(HtmlPageResponse), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    [ProducesResponseType(StatusCodes.Status404NotFound)]
    [ProducesResponseType(StatusCodes.Status500InternalServerError)]
    public IActionResult HtmlPage([FromBody] HtmlPageRequest request, [FromServices] IHtmlPageService htmlPageService)
    {
        try
        {
            var result = htmlPageService.Export(request);
            return Ok(result);
        }
        catch (ArgumentException ex)
        {
            return BadRequest(ex.Message);
        }
        catch (FileNotFoundException ex)
        {
            return NotFound($"File not found: {ex.FileName ?? ex.Message}");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "HTML page export failed.");
            return Problem(title: "HTML page export failed", detail: ex.Message, statusCode: StatusCodes.Status500InternalServerError);
        }
    }

    private static List<string> GetRequestedPaths(ExtractPptxRequest request)
    {
        var paths = new List<string>();

        if (!string.IsNullOrWhiteSpace(request.Path))
        {
            paths.Add(request.Path.Trim());
        }

        if (request.Paths is not null)
        {
            foreach (var path in request.Paths)
            {
                if (!string.IsNullOrWhiteSpace(path))
                {
                    paths.Add(path.Trim());
                }
            }
        }

        if (!string.IsNullOrWhiteSpace(request.Directory))
        {
            AddPptxFromDirectory(paths, request.Directory);
        }

        if (request.Directories is not null)
        {
            foreach (var directory in request.Directories)
            {
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    AddPptxFromDirectory(paths, directory);
                }
            }
        }

        return paths
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static void AddPptxFromDirectory(List<string> paths, string directory)
    {
        var fullDir = Path.GetFullPath(directory);
        if (!Directory.Exists(fullDir))
        {
            throw new ArgumentException($"Directory not found: {fullDir}");
        }

        foreach (var file in Directory.EnumerateFiles(fullDir, "*.pptx", SearchOption.AllDirectories))
        {
            paths.Add(file);
        }
    }

    private static bool HasDirectories(ExtractPptxRequest request)
    {
        if (!string.IsNullOrWhiteSpace(request.Directory))
        {
            return true;
        }

        return request.Directories is not null &&
            request.Directories.Any(directory => !string.IsNullOrWhiteSpace(directory));
    }

    private ExtractOptions BuildExtractOptions(ExtractPptxRequest request)
    {
        var generateScreenshots = request.GenerateScreenshots
            ?? _configuration.GetValue("LibreOffice:GenerateScreenshots", false);
        if (request.GenerateScreenshots is null && !string.IsNullOrWhiteSpace(request.SofficePath))
        {
            generateScreenshots = true;
        }
        if (request.GenerateScreenshots is null && !string.IsNullOrWhiteSpace(_configuration["LibreOffice:SofficePath"]))
        {
            generateScreenshots = true;
        }

        var screenshotMaxWidth = request.ScreenshotMaxWidth
            ?? _configuration.GetValue<int?>("Screenshots:MaxWidth");
        var screenshotMaxHeight = request.ScreenshotMaxHeight
            ?? _configuration.GetValue<int?>("Screenshots:MaxHeight");
        var screenshotJpegQuality = request.ScreenshotJpegQuality
            ?? _configuration.GetValue("Screenshots:JpegQuality", 70);
        screenshotJpegQuality = Math.Clamp(screenshotJpegQuality, 1, 100);

        var screenshotPipeline = request.ScreenshotPipeline
            ?? _configuration.GetValue("Screenshots:Pipeline", "pdf");
        var screenshotPdfDpi = request.ScreenshotPdfDpi
            ?? _configuration.GetValue<int?>("Screenshots:PdfDpi");
        if (screenshotPdfDpi is not null && screenshotPdfDpi <= 0)
        {
            screenshotPdfDpi = null;
        }
        else if (screenshotPdfDpi is not null)
        {
            screenshotPdfDpi = Math.Clamp(screenshotPdfDpi.Value, 72, 300);
        }
        var screenshotPerFrame = request.ScreenshotPerFrame
            ?? _configuration.GetValue("Screenshots:PerFrame", true);
        var screenshotFrameMaxWidth = request.ScreenshotFrameMaxWidth
            ?? _configuration.GetValue<int?>("Screenshots:FrameMaxWidth");
        var screenshotFrameMaxHeight = request.ScreenshotFrameMaxHeight
            ?? _configuration.GetValue<int?>("Screenshots:FrameMaxHeight");
        var screenshotFrameAllowUpscale = request.ScreenshotFrameAllowUpscale
            ?? _configuration.GetValue("Screenshots:FrameAllowUpscale", true);

        var screenshotParallelism = request.ScreenshotParallelism
            ?? _configuration.GetValue("Screenshots:Parallelism", 4);
        screenshotParallelism = Math.Max(1, screenshotParallelism);

        var configuredSofficePath = string.IsNullOrWhiteSpace(request.SofficePath)
            ? _configuration["LibreOffice:SofficePath"]
            : request.SofficePath;

        return new ExtractOptions(
            generateScreenshots,
            configuredSofficePath,
            screenshotMaxWidth,
            screenshotMaxHeight,
            screenshotJpegQuality,
            screenshotPipeline,
            screenshotPdfDpi,
            screenshotParallelism,
            screenshotPerFrame,
            screenshotFrameMaxWidth,
            screenshotFrameMaxHeight,
            screenshotFrameAllowUpscale);
    }

    private static List<string> GetRequestedHtmlPaths(ExportHtmlRequest request)
    {
        var paths = new List<string>();

        if (!string.IsNullOrWhiteSpace(request.Path))
        {
            paths.Add(request.Path.Trim());
        }

        if (request.Paths is not null)
        {
            foreach (var path in request.Paths)
            {
                if (!string.IsNullOrWhiteSpace(path))
                {
                    paths.Add(path.Trim());
                }
            }
        }

        return paths
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private ExtractPptxResponse ExtractSingle(string inputPath, ExtractOptions options)
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

        var fileInfo = new FileInfo(fullPath);
        if (fileInfo.Directory is null)
        {
            throw new ArgumentException("Input path does not have a valid directory.");
        }

        var outputDir = Path.Combine(fileInfo.Directory.FullName, Path.GetFileNameWithoutExtension(fileInfo.Name));
        var textsDir = Path.Combine(outputDir, "texts");
        var attachmentsDir = Path.Combine(outputDir, "attachments");
        var screenshotsDir = Path.Combine(outputDir, "screenshots");

        Directory.CreateDirectory(textsDir);
        Directory.CreateDirectory(attachmentsDir);
        Directory.CreateDirectory(screenshotsDir);

        var slideCount = 0;
        var textFilesWritten = 0;

        var attachmentFiles = new List<string>();
        var screenshotFiles = new List<string>();
        var frameMetadata = new List<FrameScreenshotInfo>();
        Dictionary<int, List<SlideFrameRect>>? framesBySlide = null;
        SlideSizeInfo? slideSizeInfo = null;

        using (var presentation = PresentationDocument.Open(fullPath, false))
        {
            var presentationPart = presentation.PresentationPart;
            var slideIds = presentationPart?.Presentation?.SlideIdList?.ChildElements.OfType<SlideId>().ToList();

            if (slideIds is not null && presentationPart is not null)
            {
                var needsSlideSize = options.ScreenshotPerFrame ||
                    (options.GenerateScreenshots &&
                     string.Equals(options.ScreenshotPipeline, "pdf", StringComparison.OrdinalIgnoreCase) &&
                     options.ScreenshotPdfDpi is null);

                if (needsSlideSize)
                {
                    slideSizeInfo = GetSlideSizeInfo(presentationPart);
                    if (options.ScreenshotPerFrame)
                    {
                        framesBySlide = ExtractSlideFrames(presentationPart, slideIds);
                    }
                }

                foreach (var slideId in slideIds)
                {
                    slideCount++;
                    var relationshipId = slideId.RelationshipId?.Value;
                    if (string.IsNullOrWhiteSpace(relationshipId))
                    {
                        continue;
                    }

                    if (presentationPart.GetPartById(relationshipId) is not SlidePart slidePart)
                    {
                        continue;
                    }

                    var lines = ExtractSlideText(slidePart);
                    var textPath = Path.Combine(textsDir, $"slide-{slideCount:000}.txt");
                    System.IO.File.WriteAllLines(textPath, lines);
                    textFilesWritten++;

                    attachmentFiles.AddRange(ExtractSlideImages(slidePart, slideCount, attachmentsDir));
                }
            }
        }

        attachmentFiles.AddRange(ExtractAttachments(fullPath, attachmentsDir));

        if (options.GenerateScreenshots)
        {
            var sofficePath = ResolveSofficePath(options.SofficePath);
            if (sofficePath is null)
            {
                throw new InvalidOperationException("LibreOffice (soffice) not found. Provide 'sofficePath', set LibreOffice:SofficePath, or set generateScreenshots=false.");
            }

            if (string.Equals(options.ScreenshotPipeline, "pdf", StringComparison.OrdinalIgnoreCase))
            {
                screenshotFiles.AddRange(ExtractSlideScreenshotsViaPdf(
                    fullPath,
                    screenshotsDir,
                    sofficePath,
                    slideCount,
                    options.ScreenshotMaxWidth,
                    options.ScreenshotMaxHeight,
                    options.ScreenshotJpegQuality,
                    options.ScreenshotPdfDpi,
                    options.ScreenshotPerFrame,
                    framesBySlide,
                    slideSizeInfo,
                    options.ScreenshotFrameMaxWidth,
                    options.ScreenshotFrameMaxHeight,
                    options.ScreenshotFrameAllowUpscale,
                    frameMetadata));
            }
            else
            {
                screenshotFiles.AddRange(ExtractSlideScreenshotsViaLibreOfficePng(
                    fullPath,
                    screenshotsDir,
                    sofficePath,
                    slideCount,
                    options.ScreenshotMaxWidth,
                    options.ScreenshotMaxHeight,
                    options.ScreenshotJpegQuality,
                    options.ScreenshotParallelism,
                    options.ScreenshotPerFrame,
                    framesBySlide,
                    slideSizeInfo,
                    options.ScreenshotFrameMaxWidth,
                    options.ScreenshotFrameMaxHeight,
                    options.ScreenshotFrameAllowUpscale,
                    frameMetadata));
            }
        }

        string? frameMetadataFile = null;
        if (frameMetadata.Count > 0 && slideSizeInfo is not null)
        {
            var frameMetadataPath = Path.Combine(screenshotsDir, "frames.json");
            var payload = new FrameMetadataFile(
                slideSizeInfo.Value.WidthEmu,
                slideSizeInfo.Value.HeightEmu,
                frameMetadata);
            var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
            {
                WriteIndented = true
            });
            System.IO.File.WriteAllText(frameMetadataPath, json);
            frameMetadataFile = frameMetadataPath;
        }

        return new ExtractPptxResponse(
            fullPath,
            outputDir,
            slideCount,
            textFilesWritten,
            attachmentFiles.Count,
            attachmentFiles,
            slideCount,
            screenshotFiles.Count,
            screenshotFiles.Count,
            screenshotFiles,
            frameMetadata.Count,
            frameMetadataFile);
    }

    private static List<string> ExtractSlideText(SlidePart slidePart)
    {
        var lines = new List<string>();

        foreach (var paragraph in slidePart.Slide.Descendants<A.Paragraph>())
        {
            var text = string.Concat(paragraph.Descendants<A.Text>().Select(t => t.Text));
            if (!string.IsNullOrWhiteSpace(text))
            {
                lines.Add(text.Trim());
            }
        }

        return lines;
    }

    private static List<string> ExtractAttachments(string pptxPath, string attachmentsDir)
    {
        var saved = new List<string>();

        using var stream = System.IO.File.OpenRead(pptxPath);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        foreach (var entry in archive.Entries)
        {
            if (string.IsNullOrEmpty(entry.Name))
            {
                continue;
            }

            if (!entry.FullName.StartsWith("ppt/embeddings/", StringComparison.OrdinalIgnoreCase) &&
                !entry.FullName.StartsWith("ppt/linkedMedia/", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var relative = entry.FullName["ppt/".Length..];
            var destinationPath = Path.Combine(attachmentsDir, relative);
            var destinationDir = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(destinationDir))
            {
                Directory.CreateDirectory(destinationDir);
            }

            entry.ExtractToFile(destinationPath, overwrite: true);
            saved.Add(destinationPath);
        }

        return saved;
    }

    internal static SlideSizeInfo GetSlideSizeInfo(PresentationPart presentationPart)
    {
        var slideSize = presentationPart.Presentation?.SlideSize;
        if (slideSize is null)
        {
            return new SlideSizeInfo(0, 0);
        }

        var width = slideSize.Cx?.Value ?? 0;
        var height = slideSize.Cy?.Value ?? 0;
        return new SlideSizeInfo(width, height);
    }

    internal static Dictionary<int, List<SlideFrameRect>> ExtractSlideFrames(
        PresentationPart presentationPart,
        List<SlideId> slideIds)
    {
        var framesBySlide = new Dictionary<int, List<SlideFrameRect>>();
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

            var frames = new List<SlideFrameRect>();

            foreach (var shape in slidePart.Slide.Descendants<Shape>())
            {
                if (TryGetFrameRect(shape.ShapeProperties?.Transform2D, out var rect))
                {
                    frames.Add(rect);
                }
            }

            foreach (var picture in slidePart.Slide.Descendants<Picture>())
            {
                if (TryGetFrameRect(picture.ShapeProperties?.Transform2D, out var rect))
                {
                    frames.Add(rect);
                }
            }

            foreach (var frame in slidePart.Slide.Descendants<GraphicFrame>())
            {
                if (TryGetFrameRect(frame.Transform, out var rect))
                {
                    frames.Add(rect);
                }
            }

            if (frames.Count > 0)
            {
                framesBySlide[slideIndex] = frames;
            }
        }

        return framesBySlide;
    }

    private static bool TryGetFrameRect(A.Transform2D? transform, out SlideFrameRect rect)
    {
        rect = default;
        if (transform?.Offset is null || transform.Extents is null)
        {
            return false;
        }

        var x = transform.Offset.X?.Value ?? 0;
        var y = transform.Offset.Y?.Value ?? 0;
        var cx = transform.Extents.Cx?.Value ?? 0;
        var cy = transform.Extents.Cy?.Value ?? 0;

        if (cx <= 0 || cy <= 0)
        {
            return false;
        }

        rect = new SlideFrameRect(x, y, cx, cy);
        return true;
    }

    private static bool TryGetFrameRect(DocumentFormat.OpenXml.Presentation.Transform? transform, out SlideFrameRect rect)
    {
        rect = default;
        if (transform?.Offset is null || transform.Extents is null)
        {
            return false;
        }

        var x = transform.Offset.X?.Value ?? 0;
        var y = transform.Offset.Y?.Value ?? 0;
        var cx = transform.Extents.Cx?.Value ?? 0;
        var cy = transform.Extents.Cy?.Value ?? 0;

        if (cx <= 0 || cy <= 0)
        {
            return false;
        }

        rect = new SlideFrameRect(x, y, cx, cy);
        return true;
    }

    private static List<string> ExtractSlideImages(SlidePart slidePart, int slideIndex, string attachmentsDir)
    {
        var saved = new List<string>();
        var imageIndex = 0;

        foreach (var blip in slidePart.Slide.Descendants<A.Blip>())
        {
            var relationshipId = blip.Embed?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId))
            {
                continue;
            }

            if (slidePart.GetPartById(relationshipId) is not ImagePart imagePart)
            {
                continue;
            }

            imageIndex++;
            var extension = GetImageExtension(imagePart);
            var fileName = $"slide-{slideIndex:000}-atachment{imageIndex}{extension}";
            var destinationPath = Path.Combine(attachmentsDir, fileName);

            using var input = imagePart.GetStream();
            using var output = System.IO.File.Create(destinationPath);
            input.CopyTo(output);

            saved.Add(destinationPath);
        }

        return saved;
    }

    internal static List<string> ExtractSlideScreenshotsViaPdf(
        string pptxPath,
        string screenshotsDir,
        string sofficePath,
        int expectedSlideCount,
        int? maxWidth,
        int? maxHeight,
        int jpegQuality,
        int? pdfDpi,
        bool perFrame,
        Dictionary<int, List<SlideFrameRect>>? framesBySlide,
        SlideSizeInfo? slideSizeInfo,
        int? frameMaxWidth,
        int? frameMaxHeight,
        bool frameAllowUpscale,
        List<FrameScreenshotInfo> frameMetadata)
    {
        var saved = new List<string>();

        var tempDir = Path.Combine(Path.GetTempPath(), $"webpptx-{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        try
        {
            var preparedPptx = PreparePptxForScreenshots(pptxPath, tempDir);
            var timeoutSeconds = Math.Max(60, expectedSlideCount * 5);
            var pdfPath = ConvertPptxToPdf(sofficePath, preparedPptx, tempDir, timeoutSeconds);

            using var pdfStream = System.IO.File.OpenRead(pdfPath);
            using var document = PdfDocument.Load(pdfStream);
            var pageCount = document.PageCount;
            var renderDpi = ComputePdfDpi(pdfDpi, slideSizeInfo, maxWidth, maxHeight);

            for (var page = 0; page < pageCount; page++)
            {
                using var rendered = document.Render(page, renderDpi, renderDpi, PdfRenderFlags.Annotations);
                using var buffer = new MemoryStream();
                rendered.Save(buffer, DrawingImageFormat.Png);
                buffer.Position = 0;

                using var image = ImageSharp.Load<Rgba32>(buffer);
                var slideIndex = page + 1;
                if (perFrame && framesBySlide is not null && slideSizeInfo is not null)
                {
                    var frameInfos = SaveFrameScreenshots(
                        image,
                        screenshotsDir,
                        slideIndex,
                        framesBySlide,
                        slideSizeInfo.Value,
                        frameMaxWidth ?? maxWidth,
                        frameMaxHeight ?? maxHeight,
                        jpegQuality,
                        frameAllowUpscale);
                    if (frameInfos.Count > 0)
                    {
                        saved.AddRange(frameInfos.Select(info => info.FilePath));
                        frameMetadata.AddRange(frameInfos);
                        continue;
                    }
                }

                using var clone = image.Clone();
                var destinationPath = Path.Combine(screenshotsDir, $"slide-{slideIndex:000}-screenshot.jpg");
                SaveScreenshotAsJpeg(clone, destinationPath, maxWidth, maxHeight, jpegQuality);
                saved.Add(destinationPath);
                AddSlideScreenshotMetadata(frameMetadata, slideIndex, destinationPath, slideSizeInfo);
            }

            return saved;
        }
        finally
        {
            try
            {
                Directory.Delete(tempDir, recursive: true);
            }
            catch
            {
            }
        }
    }

    private static List<string> ExtractSlideScreenshotsViaLibreOfficePng(
        string pptxPath,
        string screenshotsDir,
        string sofficePath,
        int expectedSlideCount,
        int? maxWidth,
        int? maxHeight,
        int jpegQuality,
        int parallelism,
        bool perFrame,
        Dictionary<int, List<SlideFrameRect>>? framesBySlide,
        SlideSizeInfo? slideSizeInfo,
        int? frameMaxWidth,
        int? frameMaxHeight,
        bool frameAllowUpscale,
        List<FrameScreenshotInfo> frameMetadata)
    {
        var saved = new List<string>();

        var tempDir = Path.Combine(Path.GetTempPath(), $"webpptx-{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        try
        {
            var preparedPptx = PreparePptxForScreenshots(pptxPath, tempDir);
            var outputDir = Path.Combine(tempDir, "out");
            var timeoutSeconds = Math.Max(60, expectedSlideCount * 10);
            var filters = new[] { "png:impress_png_Export", "png" };

            var pngFiles = ConvertWithLibreOffice(sofficePath, preparedPptx, outputDir, filters, timeoutSeconds);
            var ordered = false;
            if (pngFiles.Count < expectedSlideCount && expectedSlideCount > 0)
            {
                pngFiles = ConvertSlidesIndividually(sofficePath, preparedPptx, tempDir, filters, timeoutSeconds, expectedSlideCount, parallelism);
                ordered = true;
            }

            var baseName = Path.GetFileNameWithoutExtension(preparedPptx);
            var orderedFiles = ordered ? pngFiles : OrderConvertedSlides(pngFiles, baseName);

            var index = 0;
            foreach (var file in orderedFiles)
            {
                index++;
                using var image = ImageSharp.Load<Rgba32>(file);
                if (perFrame && framesBySlide is not null && slideSizeInfo is not null)
                {
                    var frameInfos = SaveFrameScreenshots(
                        image,
                        screenshotsDir,
                        index,
                        framesBySlide,
                        slideSizeInfo.Value,
                        frameMaxWidth ?? maxWidth,
                        frameMaxHeight ?? maxHeight,
                        jpegQuality,
                        frameAllowUpscale);
                    if (frameInfos.Count > 0)
                    {
                        saved.AddRange(frameInfos.Select(info => info.FilePath));
                        frameMetadata.AddRange(frameInfos);
                        continue;
                    }
                }

                var destinationPath = Path.Combine(screenshotsDir, $"slide-{index:000}-screenshot.jpg");
                SaveScreenshotAsJpeg(image, destinationPath, maxWidth, maxHeight, jpegQuality);
                saved.Add(destinationPath);
                AddSlideScreenshotMetadata(frameMetadata, index, destinationPath, slideSizeInfo);
            }

            return saved;
        }
        finally
        {
            try
            {
                Directory.Delete(tempDir, recursive: true);
            }
            catch
            {
            }
        }
    }

    private static List<string> ConvertWithLibreOffice(
        string sofficePath,
        string inputPath,
        string outputDir,
        string[] filters,
        int timeoutSeconds,
        string? userProfileDir = null)
    {
        List<string> pngFiles = new();

        foreach (var filter in filters)
        {
            if (Directory.Exists(outputDir))
            {
                Directory.Delete(outputDir, recursive: true);
            }
            Directory.CreateDirectory(outputDir);

            var profileArgs = string.IsNullOrWhiteSpace(userProfileDir)
                ? string.Empty
                : $" -env:UserInstallation={ToFileUri(userProfileDir)}";
            var args = $"--headless --nologo --nolockcheck --norestore{profileArgs} --convert-to \"{filter}\" --outdir \"{outputDir}\" \"{inputPath}\"";
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

            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();

            if (!process.WaitForExit(TimeSpan.FromSeconds(timeoutSeconds)))
            {
                try
                {
                    process.Kill(true);
                }
                catch
                {
                }

                throw new InvalidOperationException("LibreOffice export timed out.");
            }

            if (process.ExitCode != 0)
            {
                throw new InvalidOperationException($"LibreOffice export failed: {stderr}");
            }

            pngFiles = Directory.EnumerateFiles(outputDir, "*.*", SearchOption.AllDirectories)
                .Where(path => string.Equals(Path.GetExtension(path), ".png", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (pngFiles.Count > 0)
            {
                break;
            }
        }

        return pngFiles;
    }

    private static void SaveScreenshotAsJpeg(string sourcePath, string destinationPath, int? maxWidth, int? maxHeight, int jpegQuality)
    {
        using var input = System.IO.File.OpenRead(sourcePath);
        SaveScreenshotAsJpeg(input, destinationPath, maxWidth, maxHeight, jpegQuality);
    }

    private static void SaveScreenshotAsJpeg(Stream sourceStream, string destinationPath, int? maxWidth, int? maxHeight, int jpegQuality)
    {
        using var image = ImageSharp.Load<Rgba32>(sourceStream);
        SaveScreenshotAsJpeg(image, destinationPath, maxWidth, maxHeight, jpegQuality);
    }

    private static void SaveScreenshotAsJpeg(Image<Rgba32> image, string destinationPath, int? maxWidth, int? maxHeight, int jpegQuality)
    {
        ResizeImageIfNeeded(image, maxWidth, maxHeight, allowUpscale: false);

        var encoder = new JpegEncoder { Quality = jpegQuality };
        using var output = System.IO.File.Create(destinationPath);
        image.Save(output, encoder);
    }

    private static void ResizeImageIfNeeded(Image<Rgba32> image, int? maxWidth, int? maxHeight, bool allowUpscale)
    {
        var targetWidth = maxWidth.GetValueOrDefault(image.Width);
        var targetHeight = maxHeight.GetValueOrDefault(image.Height);

        if (maxWidth is null || maxWidth <= 0)
        {
            targetWidth = image.Width;
        }

        if (maxHeight is null || maxHeight <= 0)
        {
            targetHeight = image.Height;
        }

        var shouldResize = image.Width > targetWidth || image.Height > targetHeight;
        if (!shouldResize && allowUpscale && (image.Width < targetWidth || image.Height < targetHeight))
        {
            shouldResize = true;
        }

        if (shouldResize)
        {
            image.Mutate(context => context.Resize(new ResizeOptions
            {
                Mode = ResizeMode.Max,
                Size = new SixLabors.ImageSharp.Size(targetWidth, targetHeight)
            }));
        }
    }

    private static void AddSlideScreenshotMetadata(
        List<FrameScreenshotInfo> frameMetadata,
        int slideIndex,
        string filePath,
        SlideSizeInfo? slideSizeInfo)
    {
        if (slideSizeInfo is null)
        {
            return;
        }

        frameMetadata.Add(new FrameScreenshotInfo(
            slideIndex,
            0,
            filePath,
            0,
            0,
            slideSizeInfo.Value.WidthEmu,
            slideSizeInfo.Value.HeightEmu,
            "slide"));
    }

    private static int ComputePdfDpi(int? requestedDpi, SlideSizeInfo? slideSizeInfo, int? maxWidth, int? maxHeight)
    {
        if (requestedDpi is not null)
        {
            return requestedDpi.Value;
        }

        if (slideSizeInfo is null || slideSizeInfo.Value.WidthEmu <= 0 || slideSizeInfo.Value.HeightEmu <= 0)
        {
            return 150;
        }

        const double emuPerInch = 914400.0;
        var widthInches = slideSizeInfo.Value.WidthEmu / emuPerInch;
        var heightInches = slideSizeInfo.Value.HeightEmu / emuPerInch;

        var dpiX = maxWidth is not null && maxWidth > 0 ? maxWidth.Value / widthInches : double.PositiveInfinity;
        var dpiY = maxHeight is not null && maxHeight > 0 ? maxHeight.Value / heightInches : double.PositiveInfinity;
        var dpi = Math.Min(dpiX, dpiY);

        if (double.IsInfinity(dpi) || double.IsNaN(dpi) || dpi <= 0)
        {
            dpi = 150;
        }

        return Math.Clamp((int)Math.Round(dpi), 72, 300);
    }

    private static List<FrameScreenshotInfo> SaveFrameScreenshots(
        Image<Rgba32> image,
        string screenshotsDir,
        int slideIndex,
        Dictionary<int, List<SlideFrameRect>> framesBySlide,
        SlideSizeInfo slideSizeInfo,
        int? maxWidth,
        int? maxHeight,
        int jpegQuality,
        bool allowUpscale)
    {
        if (!framesBySlide.TryGetValue(slideIndex, out var frames) || frames.Count == 0)
        {
            return new List<FrameScreenshotInfo>();
        }

        var saved = new List<FrameScreenshotInfo>();
        var frameIndex = 0;
        foreach (var frame in frames)
        {
            if (!TryGetPixelRect(frame, image.Width, image.Height, slideSizeInfo, out var rect))
            {
                continue;
            }

            frameIndex++;
            using var crop = image.Clone(context => context.Crop(rect));
            var destinationPath = Path.Combine(screenshotsDir, $"slide-{slideIndex:000}-frame{frameIndex:00}.jpg");
            ResizeImageIfNeeded(crop, maxWidth, maxHeight, allowUpscale);
            var encoder = new JpegEncoder { Quality = jpegQuality };
            using var output = System.IO.File.Create(destinationPath);
            crop.Save(output, encoder);
            saved.Add(new FrameScreenshotInfo(
                slideIndex,
                frameIndex,
                destinationPath,
                frame.X,
                frame.Y,
                frame.Cx,
                frame.Cy,
                "frame"));
        }

        return saved;
    }

    private static bool TryGetPixelRect(
        SlideFrameRect frame,
        int imageWidth,
        int imageHeight,
        SlideSizeInfo slideSizeInfo,
        out ImageSharpRectangle rect)
    {
        rect = default;
        if (slideSizeInfo.WidthEmu <= 0 || slideSizeInfo.HeightEmu <= 0)
        {
            return false;
        }

        var scaleX = imageWidth / (double)slideSizeInfo.WidthEmu;
        var scaleY = imageHeight / (double)slideSizeInfo.HeightEmu;

        var x = (int)Math.Round(frame.X * scaleX);
        var y = (int)Math.Round(frame.Y * scaleY);
        var w = (int)Math.Round(frame.Cx * scaleX);
        var h = (int)Math.Round(frame.Cy * scaleY);

        if (w <= 0 || h <= 0)
        {
            return false;
        }

        if (x < 0)
        {
            w += x;
            x = 0;
        }
        if (y < 0)
        {
            h += y;
            y = 0;
        }

        if (x >= imageWidth || y >= imageHeight)
        {
            return false;
        }

        w = Math.Min(w, imageWidth - x);
        h = Math.Min(h, imageHeight - y);

        if (w <= 0 || h <= 0)
        {
            return false;
        }

        rect = new ImageSharpRectangle(x, y, w, h);
        return true;
    }

    private static string ConvertPptxToPdf(string sofficePath, string pptxPath, string tempDir, int timeoutSeconds)
    {
        var outputDir = Path.Combine(tempDir, "pdf");
        Directory.CreateDirectory(outputDir);

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

        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();

        if (!process.WaitForExit(TimeSpan.FromSeconds(timeoutSeconds)))
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

        var pdfPath = Directory.EnumerateFiles(outputDir, "*.*", SearchOption.AllDirectories)
            .FirstOrDefault(path => string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase));

        if (string.IsNullOrWhiteSpace(pdfPath))
        {
            throw new InvalidOperationException("LibreOffice did not produce a PDF file.");
        }

        return pdfPath;
    }

    private static List<string> ConvertSlidesIndividually(
        string sofficePath,
        string preparedPptx,
        string tempDir,
        string[] filters,
        int timeoutSeconds,
        int expectedSlideCount,
        int parallelism)
    {
        var results = new string?[expectedSlideCount];

        Parallel.For(1, expectedSlideCount + 1, new ParallelOptions { MaxDegreeOfParallelism = parallelism }, slideIndex =>
        {
            var singlePptx = PrepareSingleSlidePptx(preparedPptx, tempDir, slideIndex);
            var outputDir = Path.Combine(tempDir, $"single-{slideIndex:000}");
            var profileDir = Path.Combine(tempDir, $"profile-{slideIndex:000}");
            Directory.CreateDirectory(profileDir);

            var pngFiles = ConvertWithLibreOffice(sofficePath, singlePptx, outputDir, filters, timeoutSeconds, profileDir);
            if (pngFiles.Count == 0)
            {
                return;
            }

            results[slideIndex - 1] = pngFiles[0];
        });

        return results.Where(path => !string.IsNullOrWhiteSpace(path)).Select(path => path!).ToList();
    }

    private static string PreparePptxForScreenshots(string pptxPath, string tempDir)
    {
        var preparedPath = Path.Combine(tempDir, Path.GetFileName(pptxPath));
        System.IO.File.Copy(pptxPath, preparedPath, overwrite: true);

        using var document = PresentationDocument.Open(preparedPath, true);
        var presentation = document.PresentationPart?.Presentation;
        var slideIdList = presentation?.SlideIdList;

        if (slideIdList is not null)
        {
            foreach (var slideId in slideIdList.Elements<SlideId>())
            {
                slideId.SetAttribute(new OpenXmlAttribute("show", null, "1"));
            }

            presentation!.Save();
        }

        return preparedPath;
    }

    private static string PrepareSingleSlidePptx(string preparedPptxPath, string tempDir, int slideIndex)
    {
        var singlePath = Path.Combine(tempDir, $"single-slide-{slideIndex:000}.pptx");
        System.IO.File.Copy(preparedPptxPath, singlePath, overwrite: true);

        using var document = PresentationDocument.Open(singlePath, true);
        var presentationPart = document.PresentationPart;
        var slideIdList = presentationPart?.Presentation?.SlideIdList;

        if (presentationPart is null || slideIdList is null)
        {
            return singlePath;
        }

        var slideIds = slideIdList.Elements<SlideId>().ToList();
        for (var i = 0; i < slideIds.Count; i++)
        {
            if (i == slideIndex - 1)
            {
                continue;
            }

            var slideId = slideIds[i];
            var relationshipId = slideId.RelationshipId?.Value;
            if (!string.IsNullOrWhiteSpace(relationshipId) &&
                presentationPart.GetPartById(relationshipId) is SlidePart slidePart)
            {
                presentationPart.DeletePart(slidePart);
            }

            slideId.Remove();
        }

        presentationPart.Presentation.Save();
        return singlePath;
    }

    private static List<string> OrderConvertedSlides(List<string> files, string baseName)
    {
        if (files.Count <= 1)
        {
            return files;
        }

        var hasBase = files.Any(path =>
            string.Equals(Path.GetFileNameWithoutExtension(path), baseName, StringComparison.OrdinalIgnoreCase));

        var parsed = new List<(int? Index, string Path)>();

        foreach (var file in files)
        {
            var name = Path.GetFileNameWithoutExtension(file);
            int? index = null;

            if (string.Equals(name, baseName, StringComparison.OrdinalIgnoreCase))
            {
                index = 1;
            }
            else if (name.StartsWith(baseName, StringComparison.OrdinalIgnoreCase))
            {
                var rest = name[baseName.Length..];
                if (rest.Length > 1 && (rest[0] == '_' || rest[0] == '-') &&
                    int.TryParse(rest[1..], out var parsedIndex))
                {
                    index = hasBase ? parsedIndex + 1 : parsedIndex;
                }
            }
            else
            {
                index = TryGetTrailingNumber(name);
            }

            parsed.Add((index, file));
        }

        var withIndex = parsed.Where(item => item.Index.HasValue)
            .OrderBy(item => item.Index!.Value)
            .Select(item => item.Path);

        var withoutIndex = parsed.Where(item => !item.Index.HasValue)
            .OrderBy(item => item.Path, StringComparer.OrdinalIgnoreCase)
            .Select(item => item.Path);

        return withIndex.Concat(withoutIndex).ToList();
    }

    private static int? TryGetTrailingNumber(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return null;
        }

        var end = name.Length - 1;
        while (end >= 0 && char.IsDigit(name[end]))
        {
            end--;
        }

        if (end == name.Length - 1)
        {
            return null;
        }

        var numberText = name[(end + 1)..];
        return int.TryParse(numberText, out var number) ? number : null;
    }

    private static string ToFileUri(string directoryPath)
    {
        var normalized = directoryPath.EndsWith(Path.DirectorySeparatorChar)
            ? directoryPath
            : directoryPath + Path.DirectorySeparatorChar;
        return new Uri(normalized).AbsoluteUri;
    }

    internal static string? ResolveSofficePath(string? configuredPath)
    {
        if (!string.IsNullOrWhiteSpace(configuredPath))
        {
            var expanded = Environment.ExpandEnvironmentVariables(configuredPath);
            if (System.IO.File.Exists(expanded))
            {
                return expanded;
            }
        }

        var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (var path in pathEnv.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            var candidate = Path.Combine(path.Trim(), "soffice.exe");
            if (System.IO.File.Exists(candidate))
            {
                return candidate;
            }
        }

        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

        var knownPaths = new[]
        {
            Path.Combine(programFiles, "LibreOffice", "program", "soffice.exe"),
            Path.Combine(programFilesX86, "LibreOffice", "program", "soffice.exe")
        };

        return knownPaths.FirstOrDefault(System.IO.File.Exists);
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
}

public record ExtractPptxRequest(
    string? Path,
    bool? GenerateScreenshots = null,
    string? SofficePath = null,
    int? ScreenshotMaxWidth = null,
    int? ScreenshotMaxHeight = null,
    int? ScreenshotJpegQuality = null,
    int? ScreenshotParallelism = null,
    string? ScreenshotPipeline = null,
    int? ScreenshotPdfDpi = null,
    bool? ScreenshotPerFrame = null,
    int? ScreenshotFrameMaxWidth = null,
    int? ScreenshotFrameMaxHeight = null,
    bool? ScreenshotFrameAllowUpscale = null,
    List<string>? Paths = null,
    string? Directory = null,
    List<string>? Directories = null);

public record ExtractPptxItemResult(string InputPath, bool Success, string? Error, ExtractPptxResponse? Result);

public record ExtractPptxBatchResponse(int RequestedCount, int SucceededCount, int FailedCount, List<ExtractPptxItemResult> Items);

public record ExportHtmlRequest(
    string? Path,
    List<string>? Paths = null,
    string? SofficePath = null,
    int? MaxWidth = null,
    int? MaxHeight = null,
    int? JpegQuality = null,
    int? PdfDpi = null,
    int? FrameMaxWidth = null,
    int? FrameMaxHeight = null,
    bool? FrameAllowUpscale = null);

public record ExportHtmlResponse(
    string InputPath,
    string OutputDirectory,
    string HtmlPath,
    int SlideCount,
    int FrameCount,
    List<string> FrameFiles);

public record ExportHtmlItemResult(string InputPath, bool Success, string? Error, ExportHtmlResponse? Result);

public record ExportHtmlBatchResponse(int RequestedCount, int SucceededCount, int FailedCount, List<ExportHtmlItemResult> Items);

public record HtmlPageRequest(
    string? Path,
    string? SofficePath = null,
    int? MaxWidth = null,
    int? MaxHeight = null,
    int? JpegQuality = null,
    int? PdfDpi = null);

public record HtmlPageResponse(
    string InputPath,
    string OutputDirectory,
    string HtmlPath,
    string PdfPath,
    int PageCount,
    int ImageCount,
    List<string> PageImages);

public record HtmlPageOptions(
    string? SofficePath,
    int? MaxWidth,
    int? MaxHeight,
    int JpegQuality,
    int? PdfDpi);

public record ExtractOptions(
    bool GenerateScreenshots,
    string? SofficePath,
    int? ScreenshotMaxWidth,
    int? ScreenshotMaxHeight,
    int ScreenshotJpegQuality,
    string ScreenshotPipeline,
    int? ScreenshotPdfDpi,
    int ScreenshotParallelism,
    bool ScreenshotPerFrame,
    int? ScreenshotFrameMaxWidth,
    int? ScreenshotFrameMaxHeight,
    bool ScreenshotFrameAllowUpscale);

public record HtmlExportOptions(
    string? SofficePath,
    int? MaxWidth,
    int? MaxHeight,
    int JpegQuality,
    int? PdfDpi,
    int? FrameMaxWidth,
    int? FrameMaxHeight,
    bool FrameAllowUpscale);

public readonly record struct SlideFrameRect(long X, long Y, long Cx, long Cy);

public readonly record struct SlideSizeInfo(long WidthEmu, long HeightEmu);

public record ExtractPptxResponse(
    string InputPath,
    string OutputDirectory,
    int SlideCount,
    int TextFilesWritten,
    int AttachmentCount,
    List<string> AttachmentFiles,
    int ScreenshotExpectedCount,
    int ScreenshotExportedCount,
    int ScreenshotCount,
    List<string> ScreenshotFiles,
    int FrameMetadataCount,
    string? FrameMetadataFile);

public record FrameMetadataFile(long SlideWidthEmu, long SlideHeightEmu, List<FrameScreenshotInfo> Frames);

public record FrameScreenshotInfo(
    int SlideIndex,
    int FrameIndex,
    string FilePath,
    long X,
    long Y,
    long Cx,
    long Cy,
    string Type);

public record RebuildPptxResponse(string OutputPath, int SlideCount, int ItemCount);
