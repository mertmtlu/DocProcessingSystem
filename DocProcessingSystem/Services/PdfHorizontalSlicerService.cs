using BitMiracle.LibTiff.Classic;
using iText.IO.Image;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;

namespace DocProcessingSystem.Services
{
    public class PdfHorizontalSlicerService
    {
        /// <summary>
        /// Reads the CCITT G4 image from the source PDF, splits it into pixel-accurate column
        /// slices defined by consecutive pairs in <paramref name="cuts"/>, and writes each slice
        /// as a separate page in the output PDF.
        /// </summary>
        public void Slice(string inputPath, string outputPath, List<float> cuts)
        {
            if (cuts.Count < 2)
                throw new ArgumentException("At least two cut positions are required.");

            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(outputPath)!);

            // --- Extract raw CCITT data and page dimensions from source PDF ---
            Console.WriteLine("Extracting image from source PDF...");
            float pageWidthPt, pageHeightPt;
            int imageWidth, imageHeight;
            byte[] ccittData;

            using (var reader = new PdfReader(inputPath))
            using (var sourceDoc = new PdfDocument(reader))
            {
                var sourcePage = sourceDoc.GetFirstPage();
                pageWidthPt  = sourcePage.GetPageSize().GetWidth();
                pageHeightPt = sourcePage.GetPageSize().GetHeight();

                // Structure: Page -> /Resources/XObject/Xf1 -> /Resources/XObject/X0 (image)
                var x0 = sourcePage.GetPdfObject()
                    .GetAsDictionary(PdfName.Resources)
                    .GetAsDictionary(PdfName.XObject)
                    .GetAsStream(new PdfName("Xf1"))
                    .GetAsDictionary(PdfName.Resources)
                    .GetAsDictionary(PdfName.XObject)
                    .GetAsStream(new PdfName("X0"));

                imageWidth  = x0.GetAsNumber(PdfName.Width).IntValue();
                imageHeight = x0.GetAsNumber(PdfName.Height).IntValue();
                ccittData   = x0.GetBytes(false); // raw compressed bytes, no decoding
            }

            Console.WriteLine($"Image: {imageWidth}×{imageHeight}px  |  Page: {pageWidthPt}×{pageHeightPt}pt");

            // --- Decode full image to packed 1-bit scanlines via LibTiff ---
            Console.WriteLine("Decoding CCITT G4 image (this may take a moment)...");
            byte[] sourceTiff      = BuildCCITTG4TiffBytes(ccittData, imageWidth, imageHeight);
            byte[][] scanlines     = DecodeTiffScanlines(sourceTiff, imageHeight);

            double scaleX = (double)imageWidth / pageWidthPt;

            // --- Build output PDF, one page per slice ---
            Console.WriteLine("Building output PDF...");
            using var writer     = new PdfWriter(outputPath);
            using var targetDoc  = new PdfDocument(writer);

            for (int i = 0; i < cuts.Count - 1; i++)
            {
                float x1Pt  = cuts[i];
                float x2Pt  = cuts[i + 1];
                int   x1Px  = (int)Math.Round(x1Pt * scaleX);
                int   x2Px  = (int)Math.Round(x2Pt * scaleX);
                int   wPx   = x2Px - x1Px;
                float wPt   = x2Pt - x1Pt;

                byte[][] sliceLines = CropScanlines(scanlines, x1Px, wPx);
                byte[]   sliceTiff  = EncodeScanlinesToCCITTG4Tiff(sliceLines, wPx, imageHeight);

                var page   = targetDoc.AddNewPage(new PageSize(wPt, pageHeightPt));
                var canvas = new PdfCanvas(page);
                canvas.AddImageFittedIntoRectangle(
                    ImageDataFactory.Create(sliceTiff),
                    new Rectangle(0, 0, wPt, pageHeightPt),
                    false);
                canvas.Release();

                Console.WriteLine($"  Slice {i + 1}/{cuts.Count - 1}: {wPt:F0}pt  ({wPx}px)");
            }

            Console.WriteLine("Done.");
        }

        // ── TIFF Building ─────────────────────────────────────────────────────────

        // Wraps raw CCITT G4 bytes in a minimal, valid TIFF file (in-memory).
        private static byte[] BuildCCITTG4TiffBytes(byte[] ccittData, int width, int height)
        {
            const int numEntries = 11;
            const int ifdOffset  = 8;
            int dataOffset       = ifdOffset + 2 + numEntries * 12 + 4;

            using var ms = new MemoryStream(dataOffset + ccittData.Length);
            using var bw = new BinaryWriter(ms);

            // TIFF header — little-endian
            bw.Write((byte)0x49); bw.Write((byte)0x49); // "II"
            bw.Write((ushort)42);
            bw.Write((uint)ifdOffset);

            // IFD (tags must be in ascending order)
            bw.Write((ushort)numEntries);
            WriteIfdEntry(bw, 256, 4, 1, (uint)width);             // ImageWidth
            WriteIfdEntry(bw, 257, 4, 1, (uint)height);            // ImageLength
            WriteIfdEntry(bw, 258, 3, 1, 1u);                      // BitsPerSample = 1
            WriteIfdEntry(bw, 259, 3, 1, 4u);                      // Compression = CCITT T.6
            WriteIfdEntry(bw, 262, 3, 1, 0u);                      // PhotometricInterp = MinIsWhite
            WriteIfdEntry(bw, 266, 3, 1, 1u);                      // FillOrder = MSB first
            WriteIfdEntry(bw, 273, 4, 1, (uint)dataOffset);        // StripOffsets
            WriteIfdEntry(bw, 277, 3, 1, 1u);                      // SamplesPerPixel = 1
            WriteIfdEntry(bw, 278, 4, 1, (uint)height);            // RowsPerStrip = all
            WriteIfdEntry(bw, 279, 4, 1, (uint)ccittData.Length);  // StripByteCounts
            WriteIfdEntry(bw, 293, 4, 1, 0u);                      // T6Options = 0
            bw.Write((uint)0);                                      // next IFD = none

            bw.Write(ccittData);
            return ms.ToArray();
        }

        private static void WriteIfdEntry(BinaryWriter bw, ushort tag, ushort type, uint count, uint value)
        {
            bw.Write(tag); bw.Write(type); bw.Write(count); bw.Write(value);
        }

        // ── TIFF Decoding / Encoding via LibTiff ──────────────────────────────────

        private static byte[][] DecodeTiffScanlines(byte[] tiffBytes, int height)
        {
            var ms = new MemoryStream(tiffBytes);
            using var tiff = Tiff.ClientOpen("src", "r", ms, new StreamTiffIO());
            int scanlineSize = tiff.ScanlineSize();

            var rows = new byte[height][];
            for (int row = 0; row < height; row++)
            {
                rows[row] = new byte[scanlineSize];
                if (!tiff.ReadScanline(rows[row], row))
                    throw new InvalidOperationException($"CCITT decode failed at scanline {row}.");
            }
            return rows;
        }

        private static byte[] EncodeScanlinesToCCITTG4Tiff(byte[][] scanlines, int width, int height)
        {
            var ms = new MemoryStream();
            using (var tiff = Tiff.ClientOpen("dst", "w", ms, new StreamTiffIO()))
            {
                tiff.SetField(TiffTag.IMAGEWIDTH,    width);
                tiff.SetField(TiffTag.IMAGELENGTH,   height);
                tiff.SetField(TiffTag.SAMPLESPERPIXEL, 1);
                tiff.SetField(TiffTag.BITSPERSAMPLE, 1);
                tiff.SetField(TiffTag.COMPRESSION,   Compression.CCITTFAX4);
                tiff.SetField(TiffTag.PHOTOMETRIC,   Photometric.MINISWHITE);
                tiff.SetField(TiffTag.ROWSPERSTRIP,  height);

                for (int row = 0; row < height; row++)
                {
                    if (!tiff.WriteScanline(scanlines[row], row))
                        throw new InvalidOperationException($"CCITT encode failed at scanline {row}.");
                }
            }
            return ms.ToArray();
        }

        // ── Pixel Cropping ────────────────────────────────────────────────────────

        private static byte[][] CropScanlines(byte[][] src, int srcX, int dstWidth)
        {
            int dstBytes = (dstWidth + 7) / 8;
            var result   = new byte[src.Length][];
            for (int row = 0; row < src.Length; row++)
                result[row] = ExtractBitRange(src[row], srcX, dstWidth, dstBytes);
            return result;
        }

        // Extracts dstWidth bits starting at bit-offset srcX from a packed MSB-first byte array.
        private static byte[] ExtractBitRange(byte[] src, int srcX, int dstWidth, int dstBytes)
        {
            var dst    = new byte[dstBytes];
            int shift  = srcX & 7;
            int srcOff = srcX >> 3;

            if (shift == 0)
            {
                Buffer.BlockCopy(src, srcOff, dst, 0, dstBytes);
            }
            else
            {
                int rshift = 8 - shift;
                for (int i = 0; i < dstBytes; i++)
                {
                    dst[i] = (byte)(src[srcOff + i] << shift);
                    if (srcOff + i + 1 < src.Length)
                        dst[i] |= (byte)(src[srcOff + i + 1] >> rshift);
                }
            }

            // Clear trailing padding bits in the final byte
            int trailing = dstBytes * 8 - dstWidth;
            if (trailing > 0)
                dst[dstBytes - 1] &= (byte)(0xFF << trailing);

            return dst;
        }

        // ── LibTiff Stream Adapter ────────────────────────────────────────────────

        private sealed class StreamTiffIO : TiffStream
        {
            public override int Read(object clientData, byte[] buffer, int offset, int count) =>
                ((Stream)clientData).Read(buffer, offset, count);

            public override void Write(object clientData, byte[] buffer, int offset, int count) =>
                ((Stream)clientData).Write(buffer, offset, count);

            public override long Seek(object clientData, long offset, SeekOrigin origin) =>
                ((Stream)clientData).Seek(offset, origin);

            public override void Close(object clientData) { /* stream lifecycle managed externally */ }

            public override long Size(object clientData) => ((Stream)clientData).Length;
        }
    }
}
