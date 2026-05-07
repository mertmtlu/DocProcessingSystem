using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;

namespace DocProcessingSystem.Services
{
    public class PdfHorizontalSlicerService
    {
        /// <summary>
        /// Slices a single wide PDF page into multiple pages at the given X-axis cut positions.
        /// cuts[i] to cuts[i+1] defines each output page.
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

            var sourcePage = sourceDoc.GetFirstPage();
            float pageHeight = sourcePage.GetPageSize().GetHeight();

            for (int i = 0; i < cuts.Count - 1; i++)
            {
                float x1 = cuts[i];
                float x2 = cuts[i + 1];
                float sliceWidth = x2 - x1;

                var newPage = targetDoc.AddNewPage(new iText.Kernel.Geom.PageSize(sliceWidth, pageHeight));
                var canvas = new PdfCanvas(newPage);

                // Clip to page bounds so adjacent content doesn't bleed in
                canvas.Rectangle(0, 0, sliceWidth, pageHeight);
                canvas.Clip();
                canvas.EndPath();

                var xObject = sourcePage.CopyAsFormXObject(targetDoc);
                canvas.AddXObjectAt(xObject, -x1, 0);
                canvas.Release();

                Console.WriteLine($"  Slice {i + 1}/{cuts.Count - 1}: x={x1}–{x2} ({sliceWidth:F0}pt wide)");
            }
        }
    }
}
