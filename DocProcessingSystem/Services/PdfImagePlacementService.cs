using System;
using System.IO;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace DocProcessingSystem.Services
{
    /// <summary>
    /// Options for configuring image placement on PDF pages
    /// </summary>
    public class ImagePlacementOptions
    {
        /// <summary>
        /// Page number where the image should be placed (1-based index)
        /// Ignored when PlaceOnAllPages is true
        /// </summary>
        public int PageNumber { get; set; } = 1;

        /// <summary>
        /// Whether to place the image on all pages of the PDF
        /// </summary>
        public bool PlaceOnAllPages { get; set; } = false;

        /// <summary>
        /// Specific pages to place the image on (1-based index). If null or empty, uses PageNumber or PlaceOnAllPages
        /// </summary>
        public int[] SpecificPages { get; set; } = null;

        /// <summary>
        /// X coordinate for image placement (in points, from left edge)
        /// </summary>
        public float X { get; set; } = 0;

        /// <summary>
        /// Y coordinate for image placement (in points, from bottom edge)
        /// </summary>
        public float Y { get; set; } = 0;

        /// <summary>
        /// Desired width of the image (in points). If 0, uses original or calculated width
        /// </summary>
        public float Width { get; set; } = 0;

        /// <summary>
        /// Desired height of the image (in points). If 0, uses original or calculated height
        /// </summary>
        public float Height { get; set; } = 0;

        /// <summary>
        /// Whether to maintain the original aspect ratio when scaling
        /// </summary>
        public bool MaintainAspectRatio { get; set; } = true;

        /// <summary>
        /// Predefined position for easier placement
        /// </summary>
        public ImagePosition Position { get; set; } = ImagePosition.Custom;

        /// <summary>
        /// Margin from edges when using predefined positions (in points)
        /// </summary>
        public float Margin { get; set; } = 36; // Default 0.5 inch margin

        /// <summary>
        /// Rotation angle in degrees (0-360)
        /// </summary>
        public float Rotation { get; set; } = 0;

        /// <summary>
        /// Transparency level (0.0 = fully transparent, 1.0 = fully opaque)
        /// </summary>
        public float Opacity { get; set; } = 1.0f;

        /// <summary>
        /// Whether to place image behind text (background) or in front (foreground)
        /// </summary>
        public bool PlaceInBackground { get; set; } = false;

        /// <summary>
        /// Maximum width the image can occupy (for auto-scaling)
        /// </summary>
        public float MaxWidth { get; set; } = 0;

        /// <summary>
        /// Maximum height the image can occupy (for auto-scaling)
        /// </summary>
        public float MaxHeight { get; set; } = 0;

        /// <summary>
        /// Skip certain pages when using PlaceOnAllPages (1-based index)
        /// </summary>
        public int[] SkipPages { get; set; } = null;
    }

    /// <summary>
    /// Predefined positions for image placement
    /// </summary>
    public enum ImagePosition
    {
        Custom,
        TopLeft,
        TopCenter,
        TopRight,
        MiddleLeft,
        MiddleCenter,
        MiddleRight,
        BottomLeft,
        BottomCenter,
        BottomRight
    }

    /// <summary>
    /// Static service for placing images on PDF pages
    /// </summary>
    public static class PdfImagePlacementService
    {
        /// <summary>
        /// Adds an image to PDF pages based on the specified options
        /// </summary>
        /// <param name="inputPdfPath">Path to the input PDF file</param>
        /// <param name="outputPdfPath">Path for the output PDF file</param>
        /// <param name="imagePath">Path to the image file</param>
        /// <param name="options">Options for image placement</param>
        public static void PlaceImageOnPdf(string inputPdfPath, string outputPdfPath, string imagePath, ImagePlacementOptions options)
        {
            ValidateInputs(inputPdfPath, outputPdfPath, imagePath, options);

            using (var reader = new PdfReader(inputPdfPath))
            {
                var targetPages = GetTargetPages(options, reader.NumberOfPages);

                // Validate all target pages are within range
                var invalidPages = targetPages.Where(p => p < 1 || p > reader.NumberOfPages).ToArray();
                if (invalidPages.Any())
                {
                    throw new ArgumentOutOfRangeException(nameof(options),
                        $"Page numbers {string.Join(", ", invalidPages)} are out of range. PDF has {reader.NumberOfPages} pages.");
                }

                using (var outputStream = new FileStream(outputPdfPath, FileMode.Create))
                {
                    var stamper = new PdfStamper(reader, outputStream);

                    try
                    {
                        foreach (var pageNumber in targetPages)
                        {
                            var pageOptions = CreatePageSpecificOptions(options, pageNumber);
                            PlaceImageOnPage(stamper, imagePath, pageOptions);
                        }
                    }
                    finally
                    {
                        stamper.Close();
                    }
                }
            }

            var pageDescription = options.PlaceOnAllPages ? "all pages" :
                                 options.SpecificPages?.Length > 1 ? $"{options.SpecificPages.Length} pages" :
                                 $"page {options.PageNumber}";

            Console.WriteLine($"Image successfully placed on {pageDescription} of PDF: {outputPdfPath}");
        }

        /// <summary>
        /// Places an image on all pages of a PDF (convenience method)
        /// </summary>
        /// <param name="inputPdfPath">Path to the input PDF file</param>
        /// <param name="outputPdfPath">Path for the output PDF file</param>
        /// <param name="imagePath">Path to the image file</param>
        /// <param name="position">Position where to place the image</param>
        /// <param name="margin">Margin from edges (in points)</param>
        /// <param name="placeInBackground">Whether to place behind text</param>
        /// <param name="skipPages">Pages to skip (1-based index)</param>
        public static void PlaceImageOnAllPages(string inputPdfPath, string outputPdfPath, string imagePath,
            ImagePosition position = ImagePosition.BottomRight, float margin = 36f,
            bool placeInBackground = false, int[] skipPages = null)
        {
            var options = new ImagePlacementOptions
            {
                PlaceOnAllPages = true,
                Position = position,
                Margin = margin,
                PlaceInBackground = placeInBackground,
                SkipPages = skipPages
            };

            PlaceImageOnPdf(inputPdfPath, outputPdfPath, imagePath, options);
        }

        /// <summary>
        /// Places a watermark image on all pages of a PDF
        /// </summary>
        /// <param name="inputPdfPath">Path to the input PDF file</param>
        /// <param name="outputPdfPath">Path for the output PDF file</param>
        /// <param name="watermarkImagePath">Path to the watermark image file</param>
        /// <param name="opacity">Transparency level (0.0 to 1.0)</param>
        /// <param name="position">Position of the watermark</param>
        /// <param name="skipPages">Pages to skip watermarking</param>
        public static void AddWatermarkToAllPages(string inputPdfPath, string outputPdfPath, string watermarkImagePath,
            float opacity = 0.3f, ImagePosition position = ImagePosition.MiddleCenter, int[] skipPages = null)
        {
            var options = new ImagePlacementOptions
            {
                PlaceOnAllPages = true,
                Position = position,
                PlaceInBackground = true,
                Opacity = opacity,
                SkipPages = skipPages,
                MaintainAspectRatio = true
            };

            PlaceImageOnPdf(inputPdfPath, outputPdfPath, watermarkImagePath, options);
        }

        /// <summary>
        /// Places multiple images on different pages of a PDF
        /// </summary>
        /// <param name="inputPdfPath">Path to the input PDF file</param>
        /// <param name="outputPdfPath">Path for the output PDF file</param>
        /// <param name="imageConfigs">Array of image configurations (path and options)</param>
        public static void PlaceMultipleImagesOnPdf(string inputPdfPath, string outputPdfPath,
            (string imagePath, ImagePlacementOptions options)[] imageConfigs)
        {
            if (string.IsNullOrEmpty(inputPdfPath))
                throw new ArgumentNullException(nameof(inputPdfPath));
            if (string.IsNullOrEmpty(outputPdfPath))
                throw new ArgumentNullException(nameof(outputPdfPath));
            if (imageConfigs == null || imageConfigs.Length == 0)
                throw new ArgumentNullException(nameof(imageConfigs));

            using (var reader = new PdfReader(inputPdfPath))
            {
                using (var outputStream = new FileStream(outputPdfPath, FileMode.Create))
                {
                    var stamper = new PdfStamper(reader, outputStream);

                    try
                    {
                        foreach (var (imagePath, options) in imageConfigs)
                        {
                            ValidateImageConfig(imagePath, options, reader.NumberOfPages);
                            var targetPages = GetTargetPages(options, reader.NumberOfPages);

                            foreach (var pageNumber in targetPages)
                            {
                                var pageOptions = CreatePageSpecificOptions(options, pageNumber);
                                PlaceImageOnPage(stamper, imagePath, pageOptions);
                            }
                        }
                    }
                    finally
                    {
                        stamper.Close();
                    }
                }
            }

            var totalPlacements = imageConfigs.Sum(config => GetTargetPages(config.options, 0).Length);
            Console.WriteLine($"Successfully placed {imageConfigs.Length} images with {totalPlacements} total placements on PDF: {outputPdfPath}");
        }

        /// <summary>
        /// Gets the target pages based on the options
        /// </summary>
        private static int[] GetTargetPages(ImagePlacementOptions options, int totalPages)
        {
            if (options.SpecificPages != null && options.SpecificPages.Length > 0)
            {
                // Use specific pages, excluding any skip pages
                var pages = options.SpecificPages.ToList();
                if (options.SkipPages != null)
                {
                    pages = pages.Except(options.SkipPages).ToList();
                }
                return pages.ToArray();
            }

            if (options.PlaceOnAllPages)
            {
                // Generate all page numbers, excluding skip pages
                var allPages = Enumerable.Range(1, totalPages).ToList();
                if (options.SkipPages != null)
                {
                    allPages = allPages.Except(options.SkipPages).ToList();
                }
                return allPages.ToArray();
            }

            // Single page placement
            return new[] { options.PageNumber };
        }

        /// <summary>
        /// Creates page-specific options from the base options
        /// </summary>
        private static ImagePlacementOptions CreatePageSpecificOptions(ImagePlacementOptions options, int pageNumber)
        {
            return new ImagePlacementOptions
            {
                PageNumber = pageNumber,
                X = options.X,
                Y = options.Y,
                Width = options.Width,
                Height = options.Height,
                MaintainAspectRatio = options.MaintainAspectRatio,
                Position = options.Position,
                Margin = options.Margin,
                Rotation = options.Rotation,
                Opacity = options.Opacity,
                PlaceInBackground = options.PlaceInBackground,
                MaxWidth = options.MaxWidth,
                MaxHeight = options.MaxHeight
            };
        }

        /// <summary>
        /// Places an image on a specific page using the PDF stamper
        /// </summary>
        private static void PlaceImageOnPage(PdfStamper stamper, string imagePath, ImagePlacementOptions options)
        {
            // Load the image
            var image = Image.GetInstance(imagePath);

            // Get page dimensions
            var pageSize = stamper.Reader.GetPageSizeWithRotation(options.PageNumber);
            float pageWidth = pageSize.Width;
            float pageHeight = pageSize.Height;

            // Calculate image dimensions and position
            var (finalWidth, finalHeight) = CalculateImageDimensions(image, options, pageWidth, pageHeight);
            var (finalX, finalY) = CalculateImagePosition(options, pageWidth, pageHeight, finalWidth, finalHeight);

            // Configure image properties
            image.ScaleAbsolute(finalWidth, finalHeight);
            image.SetAbsolutePosition(finalX, finalY);

            // Apply rotation if specified
            if (options.Rotation != 0)
            {
                image.RotationDegrees = options.Rotation;
            }

            // Get the content layer (background or foreground)
            PdfContentByte contentByte;
            if (options.PlaceInBackground)
            {
                contentByte = stamper.GetUnderContent(options.PageNumber);
            }
            else
            {
                contentByte = stamper.GetOverContent(options.PageNumber);
            }

            // Apply opacity if specified
            if (options.Opacity < 1.0f)
            {
                var gState = new PdfGState { FillOpacity = options.Opacity, StrokeOpacity = options.Opacity };
                contentByte.SetGState(gState);
            }

            // Add the image to the page
            contentByte.AddImage(image);

            Console.WriteLine($"Placed image {Path.GetFileName(imagePath)} at ({finalX:F1}, {finalY:F1}) " +
                            $"with dimensions {finalWidth:F1}x{finalHeight:F1} on page {options.PageNumber}");
        }

        /// <summary>
        /// Calculates the final dimensions for the image based on options
        /// </summary>
        private static (float width, float height) CalculateImageDimensions(Image image, ImagePlacementOptions options,
            float pageWidth, float pageHeight)
        {
            float originalWidth = image.Width;
            float originalHeight = image.Height;
            float aspectRatio = originalWidth / originalHeight;

            // If no dimensions specified, use original size (with max constraints if specified)
            if (options.Width <= 0 && options.Height <= 0)
            {
                float width = originalWidth;
                float height = originalHeight;

                // Apply max constraints
                if (options.MaxWidth > 0 && width > options.MaxWidth)
                {
                    width = options.MaxWidth;
                    if (options.MaintainAspectRatio)
                        height = width / aspectRatio;
                }

                if (options.MaxHeight > 0 && height > options.MaxHeight)
                {
                    height = options.MaxHeight;
                    if (options.MaintainAspectRatio)
                        width = height * aspectRatio;
                }

                return (width, height);
            }

            // If only width is specified
            if (options.Width > 0 && options.Height <= 0)
            {
                float width = options.Width;
                float height = options.MaintainAspectRatio ? width / aspectRatio : originalHeight;
                return (width, height);
            }

            // If only height is specified
            if (options.Height > 0 && options.Width <= 0)
            {
                float height = options.Height;
                float width = options.MaintainAspectRatio ? height * aspectRatio : originalWidth;
                return (width, height);
            }

            // Both dimensions specified
            if (options.MaintainAspectRatio)
            {
                // Scale to fit within the specified dimensions while maintaining aspect ratio
                float scaleX = options.Width / originalWidth;
                float scaleY = options.Height / originalHeight;
                float scale = Math.Min(scaleX, scaleY);

                return (originalWidth * scale, originalHeight * scale);
            }
            else
            {
                return (options.Width, options.Height);
            }
        }

        /// <summary>
        /// Calculates the final position for the image based on options
        /// </summary>
        private static (float x, float y) CalculateImagePosition(ImagePlacementOptions options,
            float pageWidth, float pageHeight, float imageWidth, float imageHeight)
        {
            if (options.Position == ImagePosition.Custom)
            {
                return (options.X, options.Y);
            }

            float x, y;
            float margin = options.Margin;

            switch (options.Position)
            {
                case ImagePosition.TopLeft:
                    x = margin;
                    y = pageHeight - imageHeight - margin;
                    break;

                case ImagePosition.TopCenter:
                    x = (pageWidth - imageWidth) / 2;
                    y = pageHeight - imageHeight - margin;
                    break;

                case ImagePosition.TopRight:
                    x = pageWidth - imageWidth - margin;
                    y = pageHeight - imageHeight - margin;
                    break;

                case ImagePosition.MiddleLeft:
                    x = margin;
                    y = (pageHeight - imageHeight) / 2;
                    break;

                case ImagePosition.MiddleCenter:
                    x = (pageWidth - imageWidth) / 2;
                    y = (pageHeight - imageHeight) / 2;
                    break;

                case ImagePosition.MiddleRight:
                    x = pageWidth - imageWidth - margin;
                    y = (pageHeight - imageHeight) / 2;
                    break;

                case ImagePosition.BottomLeft:
                    x = margin;
                    y = margin;
                    break;

                case ImagePosition.BottomCenter:
                    x = (pageWidth - imageWidth) / 2;
                    y = margin;
                    break;

                case ImagePosition.BottomRight:
                    x = pageWidth - imageWidth - margin;
                    y = margin;
                    break;

                default:
                    x = options.X;
                    y = options.Y;
                    break;
            }

            return (x, y);
        }

        /// <summary>
        /// Validates input parameters
        /// </summary>
        private static void ValidateInputs(string inputPdfPath, string outputPdfPath, string imagePath, ImagePlacementOptions options)
        {
            if (string.IsNullOrEmpty(inputPdfPath))
                throw new ArgumentNullException(nameof(inputPdfPath));
            if (string.IsNullOrEmpty(outputPdfPath))
                throw new ArgumentNullException(nameof(outputPdfPath));
            if (string.IsNullOrEmpty(imagePath))
                throw new ArgumentNullException(nameof(imagePath));
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            if (!File.Exists(inputPdfPath))
                throw new FileNotFoundException($"Input PDF file not found: {inputPdfPath}");
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            if (options.Opacity < 0 || options.Opacity > 1)
                throw new ArgumentOutOfRangeException(nameof(options.Opacity), "Opacity must be between 0.0 and 1.0");
        }

        /// <summary>
        /// Validates individual image configuration
        /// </summary>
        private static void ValidateImageConfig(string imagePath, ImagePlacementOptions options, int totalPages)
        {
            if (string.IsNullOrEmpty(imagePath))
                throw new ArgumentNullException(nameof(imagePath));
            if (options == null)
                throw new ArgumentNullException(nameof(options));
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            var targetPages = GetTargetPages(options, totalPages);
            var invalidPages = targetPages.Where(p => p < 1 || p > totalPages).ToArray();
            if (invalidPages.Any())
            {
                throw new ArgumentOutOfRangeException(nameof(options),
                    $"Page numbers {string.Join(", ", invalidPages)} are out of range. PDF has {totalPages} pages.");
            }
        }

        /// <summary>
        /// Helper method to convert inches to points (1 inch = 72 points)
        /// </summary>
        public static float InchesToPoints(float inches)
        {
            return inches * 72f;
        }

        /// <summary>
        /// Helper method to convert millimeters to points (1 mm = 2.834645669 points)
        /// </summary>
        public static float MillimetersToPoints(float millimeters)
        {
            return millimeters * 2.834645669f;
        }

        /// <summary>
        /// Helper method to get page dimensions in points
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <param name="pageNumber">Page number (1-based)</param>
        /// <returns>Tuple containing width and height in points</returns>
        public static (float width, float height) GetPageDimensions(string pdfPath, int pageNumber = 1)
        {
            using (var reader = new PdfReader(pdfPath))
            {
                if (pageNumber < 1 || pageNumber > reader.NumberOfPages)
                {
                    throw new ArgumentOutOfRangeException(nameof(pageNumber),
                        $"Page number {pageNumber} is out of range. PDF has {reader.NumberOfPages} pages.");
                }

                var pageSize = reader.GetPageSizeWithRotation(pageNumber);
                return (pageSize.Width, pageSize.Height);
            }
        }
    }
}