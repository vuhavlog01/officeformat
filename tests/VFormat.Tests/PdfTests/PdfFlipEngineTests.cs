using System;
using System.IO;
using System.Linq;
using Xunit;
using PdfSharp.Pdf;
using VFormat.Core.Pdf;

namespace VFormat.Tests.PdfTests
{
    public class PdfFlipEngineTests
    {
        [Fact]
        public void ParsePageOrder_DiscreteAndRanges_ReturnsExpectedIndices()
        {
            // "1, 3, 2" -> [0, 2, 1]
            var res1 = PdfFlipEngine.ParsePageOrder("1, 3, 2", 5);
            Assert.Equal(new[] { 0, 2, 1 }, res1.ToArray());

            // "1-3, 5-4" -> [0, 1, 2, 4, 3]
            var res2 = PdfFlipEngine.ParsePageOrder("1-3, 5-4", 5);
            Assert.Equal(new[] { 0, 1, 2, 4, 3 }, res2.ToArray());

            // "5-1" descending -> [4, 3, 2, 1, 0]
            var res3 = PdfFlipEngine.ParsePageOrder("5-1", 5);
            Assert.Equal(new[] { 4, 3, 2, 1, 0 }, res3.ToArray());
        }

        [Fact]
        public void ParsePageOrder_InvalidInputs_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>(() => PdfFlipEngine.ParsePageOrder("", 5));
            Assert.Throws<ArgumentException>(() => PdfFlipEngine.ParsePageOrder("   ", 5));
            Assert.Throws<ArgumentException>(() => PdfFlipEngine.ParsePageOrder("6", 5)); // Page > total
            Assert.Throws<ArgumentException>(() => PdfFlipEngine.ParsePageOrder("0", 5)); // Page < 1
            Assert.Throws<ArgumentException>(() => PdfFlipEngine.ParsePageOrder("1-6", 5)); // End page > total
            Assert.Throws<ArgumentException>(() => PdfFlipEngine.ParsePageOrder("abc", 5)); // Not a number
        }

        [Fact]
        public void UniquePath_WhenNotExists_ReturnsOriginal()
        {
            string nonExistent = Path.Combine(Path.GetTempPath(), $"unique_test_{Guid.NewGuid():N}.pdf");
            Assert.Equal(nonExistent, PdfFlipEngine.UniquePath(nonExistent));
        }

        [Fact]
        public void UniquePath_WhenExists_AppendsCounter()
        {
            string tempFile = Path.Combine(Path.GetTempPath(), $"unique_test_{Guid.NewGuid():N}.pdf");
            File.WriteAllText(tempFile, "dummy");
            try
            {
                string unique = PdfFlipEngine.UniquePath(tempFile);
                Assert.NotEqual(tempFile, unique);
                Assert.Contains("(1).pdf", unique);
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }

        [Fact]
        public void ProcessPdfNative_Reverses3Pages_CreatesValidReversedPdf()
        {
            string samplePdf = Path.Combine(Path.GetTempPath(), $"sample_test_{Guid.NewGuid():N}.pdf");
            try
            {
                using (var doc = new PdfDocument())
                {
                    doc.AddPage();
                    doc.AddPage();
                    doc.AddPage();
                    doc.Save(samplePdf);
                }

                Assert.Equal(3, PdfFlipEngine.GetPageCount(samplePdf));

                var res = PdfFlipEngine.ProcessPdfNative(
                    samplePdf,
                    outDir: Path.GetTempPath(),
                    suffix: "_reversed_test",
                    inplace: false,
                    overwrite: true,
                    reverseOrder: true);

                Assert.True(res.Success);
                Assert.Equal(3, res.OriginalPageCount);
                Assert.Equal(3, res.FinalPageCount);
                Assert.True(File.Exists(res.OutputPath));

                if (File.Exists(res.OutputPath))
                {
                    File.Delete(res.OutputPath);
                }
            }
            finally
            {
                if (File.Exists(samplePdf)) File.Delete(samplePdf);
            }
        }
    }
}
