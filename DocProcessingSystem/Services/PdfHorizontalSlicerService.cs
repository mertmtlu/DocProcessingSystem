using iText.Kernel.Geom;
using iText.Kernel.Pdf;

namespace DocProcessingSystem.Services
{
    public class PdfHorizontalSlicerService
    {
        /// <summary>
        /// Slices a single wide PDF page into multiple pages at the given X-axis cut positions.
        /// cuts[i] to cuts[i+1] defines each output page. Uses CropBox to define each slice.
        /// </summary>
        public void Slice(string inputPath, string outputPath, List<float> cuts)
        {
            if (cuts.Count < 2)
                throw new ArgumentException("At least two cut positions are required.");

            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(outputPath)!);

            using var reader = new PdfReader(inputPath);
            using var sourceDoc = new PdfDocument(reader);
            using var writer = new PdfWriter(outputPath);
            using var targetDoc = new PdfDocument(writer);

            float pageHeight = sourceDoc.GetFirstPage().GetPageSize().GetHeight();

            for (int i = 0; i < cuts.Count - 1; i++)
            {
                float x1 = cuts[i];
                float x2 = cuts[i + 1];

                sourceDoc.CopyPagesTo(1, 1, targetDoc);
                var page = targetDoc.GetLastPage();

                var cropBox = new Rectangle(x1, 0, x2 - x1, pageHeight);
                page.SetCropBox(cropBox);

                Console.WriteLine($"  Slice {i + 1}/{cuts.Count - 1}: x={x1}–{x2} ({x2 - x1:F0}pt wide)");
            }
        }
    }
}
