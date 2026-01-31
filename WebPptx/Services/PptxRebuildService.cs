using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using WebPptx.Controllers;

namespace WebPptx.Services;

public interface IPptxRebuildService
{
    RebuildPptxResult RebuildFromFrames(RebuildPptxRequest request);
}

public class PptxRebuildService : IPptxRebuildService
{
    public RebuildPptxResult RebuildFromFrames(RebuildPptxRequest request)
    {
        if (request is null || string.IsNullOrWhiteSpace(request.FramesJsonPath))
        {
            throw new ArgumentException("FramesJsonPath must be provided.");
        }

        var metadataPath = Path.GetFullPath(request.FramesJsonPath);
        if (!File.Exists(metadataPath))
        {
            throw new FileNotFoundException("frames.json not found.", metadataPath);
        }

        var metadata = LoadMetadata(metadataPath);

        var outputPath = string.IsNullOrWhiteSpace(request.OutputPath)
            ? Path.Combine(Path.GetDirectoryName(metadataPath) ?? Environment.CurrentDirectory, "rebuilt.pptx")
            : Path.GetFullPath(request.OutputPath);

        if (File.Exists(outputPath))
        {
            if (!request.Overwrite)
            {
                throw new InvalidOperationException($"Output file already exists: {outputPath}");
            }
            File.Delete(outputPath);
        }

        var frames = metadata.Frames ?? new List<FrameScreenshotInfo>();
        var slideCount = frames.Count > 0 ? frames.Max(frame => frame.SlideIndex) : 0;

        using (var document = PresentationDocument.Create(outputPath, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            var slideMasterPart = CreateSlideMasterPart(presentationPart);
            var slideLayoutPart = slideMasterPart.SlideLayoutParts.First();

            presentationPart.Presentation.SlideIdList = new SlideIdList();
            presentationPart.Presentation.SlideSize = new SlideSize
            {
                Cx = (int)metadata.SlideWidthEmu,
                Cy = (int)metadata.SlideHeightEmu,
                Type = SlideSizeValues.Custom
            };

            var slideMap = new Dictionary<int, SlidePart>();
            for (var slideIndex = 1; slideIndex <= slideCount; slideIndex++)
            {
                var slidePart = presentationPart.AddNewPart<SlidePart>();
                slidePart.Slide = CreateBlankSlide();
                slidePart.AddPart(slideLayoutPart);
                slidePart.Slide.Save();

                var slideId = presentationPart.Presentation.SlideIdList.AppendChild(new SlideId
                {
                    Id = (uint)(256 + slideIndex),
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                });

                slideMap[slideIndex] = slidePart;
            }

            foreach (var group in frames.GroupBy(frame => frame.SlideIndex))
            {
                if (!slideMap.TryGetValue(group.Key, out var slidePart))
                {
                    continue;
                }

                var frameItems = group.Where(item => string.Equals(item.Type, "frame", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(item => item.FrameIndex)
                    .ToList();
                var slideItems = group.Where(item => string.Equals(item.Type, "slide", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(item => item.FrameIndex)
                    .ToList();

                var itemsToPlace = frameItems.Count > 0 ? frameItems : (request.UseSlideFallback ? slideItems : new List<FrameScreenshotInfo>());
                foreach (var item in itemsToPlace)
                {
                    var filePath = ResolveImagePath(item.FilePath, metadataPath);
                    if (!File.Exists(filePath))
                    {
                        continue;
                    }

                    AddImageToSlide(slidePart, filePath, item.X, item.Y, item.Cx, item.Cy);
                }
            }

            presentationPart.Presentation.Save();
        }

        return new RebuildPptxResult(outputPath, slideCount, frames.Count);
    }

    private static FrameMetadataFile LoadMetadata(string metadataPath)
    {
        var json = File.ReadAllText(metadataPath);
        var metadata = JsonSerializer.Deserialize<FrameMetadataFile>(json, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        });

        if (metadata is null)
        {
            throw new InvalidOperationException("Failed to parse frames.json.");
        }

        return metadata;
    }

    private static string ResolveImagePath(string filePath, string metadataPath)
    {
        if (Path.IsPathRooted(filePath))
        {
            return filePath;
        }

        var baseDir = Path.GetDirectoryName(metadataPath) ?? Environment.CurrentDirectory;
        return Path.GetFullPath(Path.Combine(baseDir, filePath));
    }

    private static SlideMasterPart CreateSlideMasterPart(PresentationPart presentationPart)
    {
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();

        slideLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1U, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup()))),
            new ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = SlideLayoutValues.Blank
        };

        slideLayoutPart.SlideLayout.Save();

        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1U, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup()))),
            new SlideLayoutIdList(
                new SlideLayoutId { Id = 1U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) }),
            new ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            });

        slideMasterPart.SlideMaster.Save();

        presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList(
            new SlideMasterId { Id = 1U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) });

        return slideMasterPart;
    }

    private static Slide CreateBlankSlide()
    {
        return new Slide(new CommonSlideData(new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 1U, Name = "" },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()))));
    }

    private static void AddImageToSlide(SlidePart slidePart, string imagePath, long x, long y, long cx, long cy)
    {
        var imagePart = slidePart.AddNewPart<ImagePart>(GetImageContentType(imagePath));

        using (var stream = File.OpenRead(imagePath))
        {
            imagePart.FeedData(stream);
        }

        var relId = slidePart.GetIdOfPart(imagePart);
        var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;
        var shapeId = GetNextShapeId(shapeTree);

        var picture = new Picture(
            new NonVisualPictureProperties(
                new NonVisualDrawingProperties { Id = shapeId, Name = Path.GetFileName(imagePath) },
                new NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new ApplicationNonVisualDrawingProperties()),
            new BlipFill(
                new A.Blip { Embed = relId },
                new A.Stretch(new A.FillRectangle())),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        shapeTree.AppendChild(picture);
    }

    private static uint GetNextShapeId(ShapeTree shapeTree)
    {
        var maxId = shapeTree.Descendants<NonVisualDrawingProperties>()
            .Select(prop => prop.Id?.Value ?? 0U)
            .DefaultIfEmpty(0U)
            .Max();
        return maxId + 1;
    }

    private static string GetImageContentType(string imagePath)
    {
        var extension = Path.GetExtension(imagePath).ToLowerInvariant();
        return extension switch
        {
            ".png" => "image/png",
            ".jpg" => "image/jpeg",
            ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            ".tif" => "image/tiff",
            ".tiff" => "image/tiff",
            ".emf" => "image/x-emf",
            ".wmf" => "image/x-wmf",
            _ => "image/png"
        };
    }
}

public record RebuildPptxRequest(string FramesJsonPath, string? OutputPath = null, bool Overwrite = false, bool UseSlideFallback = true);

public record RebuildPptxResult(string OutputPath, int SlideCount, int ItemCount);
