using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace DocToPdfTool.Utils
{
    public class PdfToExcelConverter : IDisposable
    {
        private bool _disposed;

        private const double ColumnClusterTolerance = 6.0;
        private const double RowToleranceFactor = 0.5;
        private const double MinRowTolerance = 3.0;
        private const double ColumnAppearanceThreshold = 0.35;
        private const double MinColumnWidth = 15.0;
        private const double HeaderFooterYThreshold = 0.09;

        public void Convert(string pdfPath, string outputDir)
        {
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException("PDF文件不存在", pdfPath);

            string fileName = Path.GetFileNameWithoutExtension(pdfPath);
            string outputPath = Path.Combine(outputDir, $"{fileName}.xlsx");

            // If file already exists, append (1), (2), etc.
            if (File.Exists(outputPath))
            {
                int counter = 1;
                do
                {
                    outputPath = Path.Combine(outputDir, $"{fileName}({counter}).xlsx");
                    counter++;
                } while (File.Exists(outputPath));
            }

            var pagesData = ExtractAllPages(pdfPath);
            WriteToExcel(pagesData, outputPath);
        }

        private static List<PageResult> ExtractAllPages(string pdfPath)
        {
            var result = new List<PageResult>();

            using (var pdf = PdfDocument.Open(pdfPath))
            {
                foreach (var page in pdf.GetPages())
                {
                    try
                    {
                        var words = page.GetWords().ToList();
                        double pageHeight = page.Height;
                        var rows = ProcessPage(words, pageHeight);
                        result.Add(new PageResult(page.Number, rows));
                    }
                    catch
                    {
                        result.Add(new PageResult(page.Number, new List<List<string>>()));
                    }
                }
            }

            return result;
        }

        private static List<List<string>> ProcessPage(List<Word> words, double pageHeight)
        {
            if (words.Count == 0)
                return new List<List<string>>();

            // Filter headers and footers
            var contentWords = FilterMargins(words, pageHeight);

            if (contentWords.Count == 0)
                return new List<List<string>>();

            // Cluster into initial rows
            var rows = ClusterRows(contentWords);
            if (rows.Count == 0)
                return new List<List<string>>();

            // Detect if page has table structure
            if (IsTablePage(rows))
            {
                // Merge split rows (text wrapping within cells)
                double medianHeight = rows
                    .SelectMany(r => r)
                    .Select(w => w.BoundingBox.Height)
                    .OrderBy(h => h)
                    .ElementAt(rows.Count / 2);
                rows = MergeSplitRows(rows, medianHeight);

                // Detect columns by left-edge alignment
                var boundaries = DetectColumnsByAlignment(rows);

                if (boundaries.Count >= 2)
                    return BuildTable(rows, boundaries);
            }

            // Non-table fallback: output as single-column text lines
            return ExtractTextAsLines(rows);
        }

        /// <summary>
        /// Filter out words in the top and bottom margin (headers/footers).
        /// </summary>
        private static List<Word> FilterMargins(List<Word> words, double pageHeight)
        {
            double topThreshold = pageHeight * (1.0 - HeaderFooterYThreshold);
            double bottomThreshold = pageHeight * HeaderFooterYThreshold;

            return words.Where(w =>
            {
                double centerY = (w.BoundingBox.Top + w.BoundingBox.Bottom) / 2;
                return centerY <= topThreshold && centerY >= bottomThreshold;
            }).ToList();
        }

        /// <summary>
        /// Cluster words into rows by vertical center, with adaptive tolerance.
        /// </summary>
        private static List<List<Word>> ClusterRows(List<Word> words)
        {
            var sorted = words
                .OrderByDescending(w => (w.BoundingBox.Top + w.BoundingBox.Bottom) / 2)
                .ToList();

            double medianHeight = words
                .Select(w => w.BoundingBox.Height)
                .OrderBy(h => h)
                .ElementAt(words.Count / 2);
            double tolerance = Math.Max(medianHeight * RowToleranceFactor, MinRowTolerance);

            var rows = new List<List<Word>>();
            var currentRow = new List<Word> { sorted[0] };
            double currentY = (sorted[0].BoundingBox.Top + sorted[0].BoundingBox.Bottom) / 2;

            for (int i = 1; i < sorted.Count; i++)
            {
                var word = sorted[i];
                double wordY = (word.BoundingBox.Top + word.BoundingBox.Bottom) / 2;

                if (Math.Abs(wordY - currentY) <= tolerance)
                {
                    currentRow.Add(word);
                }
                else
                {
                    rows.Add(currentRow
                        .OrderBy(w => w.BoundingBox.Left)
                        .ToList());
                    currentRow = new List<Word> { word };
                    currentY = wordY;
                }
            }

            if (currentRow.Count > 0)
                rows.Add(currentRow
                    .OrderBy(w => w.BoundingBox.Left)
                    .ToList());

            return rows;
        }

        /// <summary>
        /// Merge adjacent rows where text is split across lines (text wrap within cells).
        /// Merges when the vertical gap between rows is &lt; 2× median word height
        /// and the X ranges of the rows overlap.
        /// </summary>
        private static List<List<Word>> MergeSplitRows(List<List<Word>> rows, double medianHeight)
        {
            if (rows.Count <= 1)
                return rows;

            double mergeThreshold = medianHeight * 2.2;
            var merged = new List<List<Word>>();
            merged.Add(rows[0]);

            for (int i = 1; i < rows.Count; i++)
            {
                var prevRow = merged[merged.Count - 1];
                var currRow = rows[i];

                double prevY = prevRow.Min(w => (w.BoundingBox.Top + w.BoundingBox.Bottom) / 2);
                double currY = currRow.Max(w => (w.BoundingBox.Top + w.BoundingBox.Bottom) / 2);
                double gap = prevY - currY;

                // Check if X ranges overlap
                double prevMinX = prevRow.Min(w => w.BoundingBox.Left);
                double prevMaxX = prevRow.Max(w => w.BoundingBox.Right);
                double currMinX = currRow.Min(w => w.BoundingBox.Left);
                double currMaxX = currRow.Max(w => w.BoundingBox.Right);
                bool xOverlap = prevMinX < currMaxX && currMinX < prevMaxX;

                if (gap <= mergeThreshold && xOverlap)
                {
                    // Merge: add current row's words to previous row
                    prevRow.AddRange(currRow);
                }
                else
                {
                    merged.Add(currRow);
                }
            }

            // Re-sort words within each merged row
            for (int i = 0; i < merged.Count; i++)
            {
                merged[i] = merged[i]
                    .OrderBy(w => w.BoundingBox.Left)
                    .ToList();
            }

            return merged;
        }

        /// <summary>
        /// Detect if a set of rows forms a table.
        /// A table page has many rows where words form distinct column-like clusters.
        /// </summary>
        private static bool IsTablePage(List<List<Word>> rows)
        {
            if (rows.Count < 3)
                return false;

            int multiColumnCount = 0;
            int totalContentRows = 0;

            foreach (var row in rows)
            {
                if (row.Count < 2)
                    continue;

                totalContentRows++;

                var sorted = row.OrderBy(w => w.BoundingBox.Left).ToList();
                int clusters = 1;
                for (int i = 1; i < sorted.Count; i++)
                {
                    double gap = sorted[i].BoundingBox.Left - sorted[i - 1].BoundingBox.Right;
                    if (gap > 15)
                        clusters++;
                }

                if (clusters >= 3)
                    multiColumnCount++;
            }

            if (totalContentRows == 0)
                return false;

            return (double)multiColumnCount / totalContentRows >= 0.25;
        }

        /// <summary>
        /// Detect column boundaries by left-edge alignment.
        /// Only uses rows that look like table rows (3+ word clusters) for detection,
        /// then filters out spurious narrow columns.
        /// </summary>
        private static List<double> DetectColumnsByAlignment(List<List<Word>> rows)
        {
            // Identify table-like rows (rows with 3+ word clusters)
            var tableRows = rows.Where(row =>
            {
                if (row.Count < 3) return false;
                var sorted = row.OrderBy(w => w.BoundingBox.Left).ToList();
                int clusters = 1;
                for (int i = 1; i < sorted.Count; i++)
                    if (sorted[i].BoundingBox.Left - sorted[i - 1].BoundingBox.Right > 15)
                        clusters++;
                return clusters >= 3;
            }).ToList();

            // Fall back to using all rows with 2+ words if no 3-cluster rows found
            if (tableRows.Count < 3)
            {
                tableRows = rows.Where(r => r.Count >= 2).ToList();
            }

            if (tableRows.Count < 2)
                return new List<double>();

            // Collect all left edges from table rows only
            var allLeftEdges = new List<double>();
            foreach (var row in tableRows)
            {
                foreach (var word in row)
                {
                    allLeftEdges.Add(Math.Round(word.BoundingBox.Left, 1));
                }
            }

            if (allLeftEdges.Count == 0)
                return new List<double>();

            // Cluster by proximity
            allLeftEdges.Sort();
            var edgeClusters = new List<List<double>>();
            var currentCluster = new List<double> { allLeftEdges[0] };
            for (int i = 1; i < allLeftEdges.Count; i++)
            {
                if (allLeftEdges[i] - currentCluster.Last() <= ColumnClusterTolerance)
                    currentCluster.Add(allLeftEdges[i]);
                else
                {
                    edgeClusters.Add(currentCluster);
                    currentCluster = new List<double> { allLeftEdges[i] };
                }
            }
            edgeClusters.Add(currentCluster);

            // For each cluster, check how many rows have a word starting at this X
            var columnStarts = new List<double>();
            int minAppearances = (int)Math.Ceiling(tableRows.Count * ColumnAppearanceThreshold);

            foreach (var cluster in edgeClusters)
            {
                if (cluster.Count < minAppearances)
                    continue;

                double avgX = cluster.Average();
                int rowCount = tableRows.Count(row =>
                    row.Any(w => Math.Abs(Math.Round(w.BoundingBox.Left, 1) - avgX) <= ColumnClusterTolerance));

                if (rowCount >= minAppearances)
                    columnStarts.Add(avgX);
            }

            if (columnStarts.Count < 2)
                return new List<double>();

            columnStarts.Sort();

            // Filter out spurious narrow columns (< MinColumnWidth apart)
            var filtered = new List<double> { columnStarts[0] };
            for (int i = 1; i < columnStarts.Count; i++)
            {
                if (columnStarts[i] - filtered.Last() >= MinColumnWidth)
                    filtered.Add(columnStarts[i]);
            }

            if (filtered.Count < 2)
                return new List<double>();

            // Build boundaries: left edge of each column, and right edge of last column
            var boundaries = new List<double>();
            for (int i = 0; i < filtered.Count; i++)
            {
                boundaries.Add(filtered[i]);
            }

            // Add right edge of last column
            double maxRight = tableRows
                .SelectMany(r => r)
                .Where(w => Math.Abs(Math.Round(w.BoundingBox.Left, 1) - filtered.Last()) <= ColumnClusterTolerance)
                .Max(w => w.BoundingBox.Right);
            boundaries.Add(maxRight + 2);

            return boundaries;
        }

        /// <summary>
        /// Assign words to cells based on column boundaries using center-based assignment.
        /// Each word is assigned to the column whose center X it falls within.
        /// </summary>
        private static List<List<string>> BuildTable(List<List<Word>> rows, List<double> boundaries)
        {
            // Compute column centers
            var colCenters = new double[boundaries.Count - 1];
            for (int c = 0; c < colCenters.Length; c++)
                colCenters[c] = (boundaries[c] + boundaries[c + 1]) / 2;

            var table = new List<List<string>>();

            foreach (var row in rows)
            {
                var cells = new string[colCenters.Length];
                for (int ci = 0; ci < colCenters.Length; ci++)
                    cells[ci] = "";

                foreach (var word in row)
                {
                    double wordCenter = (word.BoundingBox.Left + word.BoundingBox.Right) / 2;

                    // Find the closest column center
                    int bestCol = 0;
                    double bestDist = double.MaxValue;
                    for (int c = 0; c < colCenters.Length; c++)
                    {
                        double dist = Math.Abs(wordCenter - colCenters[c]);
                        if (dist < bestDist)
                        {
                            bestDist = dist;
                            bestCol = c;
                        }
                    }

                    // Only assign if the word center is within the column boundaries
                    if (wordCenter >= boundaries[bestCol] && wordCenter <= boundaries[bestCol + 1])
                    {
                        cells[bestCol] += word.Text.Trim();
                    }
                }

                table.Add(cells.ToList());
            }

            return table;
        }

        /// <summary>
        /// For non-table pages: output each visual line as a single cell.
        /// </summary>
        private static List<List<string>> ExtractTextAsLines(List<List<Word>> rows)
        {
            var result = new List<List<string>>();

            foreach (var row in rows)
            {
                string line = string.Join(" ", row
                    .OrderBy(w => w.BoundingBox.Left)
                    .Select(w => w.Text.Trim()));

                if (!string.IsNullOrWhiteSpace(line))
                    result.Add(new List<string> { line.Trim() });
            }

            return result;
        }

        // ===== XLSX Writer =====

        private static void WriteToExcel(List<PageResult> pagesData, string outputPath)
        {
            var ns = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var relNs = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            var ctNs = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types");
            var relPkgNs = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships");

            using (var archive = ZipFile.Open(outputPath, ZipArchiveMode.Create))
            {
                var contentTypes = new XElement(ctNs + "Types",
                    new XElement(ctNs + "Default",
                        new XAttribute("Extension", "rels"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                    new XElement(ctNs + "Default",
                        new XAttribute("Extension", "xml"),
                        new XAttribute("ContentType", "application/xml"))
                );

                var rels = new XElement(relPkgNs + "Relationships",
                    new XElement(relPkgNs + "Relationship",
                        new XAttribute("Id", "rId1"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
                        new XAttribute("Target", "xl/workbook.xml"))
                );

                var workbook = new XElement(ns + "workbook",
                    new XAttribute(XNamespace.Xmlns + "r", relNs.NamespaceName),
                    new XElement(ns + "sheets"));

                var wbRels = new XElement(relPkgNs + "Relationships");
                uint sheetId = 0;

                foreach (var pageData in pagesData)
                {
                    sheetId++;
                    string sheetFileName = $"worksheets/sheet{sheetId}.xml";
                    string relId = $"rId{sheetId}";

                    workbook.Element(ns + "sheets").Add(new XElement(ns + "sheet",
                        new XAttribute("name", pagesData.Count > 1 ? $"第{pageData.PageNumber}页" : "Sheet1"),
                        new XAttribute("sheetId", sheetId),
                        new XAttribute(relNs + "id", relId)));

                    wbRels.Add(new XElement(relPkgNs + "Relationship",
                        new XAttribute("Id", relId),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                        new XAttribute("Target", sheetFileName)));

                    // Build sheet XML
                    var sheetData = new XElement(ns + "sheetData");
                    uint rowIdx = 0;

                    foreach (var row in pageData.Rows)
                    {
                        rowIdx++;
                        var rowElem = new XElement(ns + "row", new XAttribute("r", rowIdx));
                        uint colIdx = 0;

                        foreach (var cellText in row)
                        {
                            colIdx++;
                            string colRef = GetColumnLetter(colIdx) + rowIdx;

                            var cell = new XElement(ns + "c",
                                new XAttribute("r", colRef),
                                new XAttribute("t", "inlineStr"),
                                new XElement(ns + "is",
                                    new XElement(ns + "t", cellText ?? "")));
                            rowElem.Add(cell);
                        }

                        sheetData.Add(rowElem);
                    }

                    var worksheet = new XElement(ns + "worksheet", sheetData);
                    var sheetEntry = archive.CreateEntry($"xl/{sheetFileName}", CompressionLevel.Optimal);
                    using (var sw = new StreamWriter(sheetEntry.Open(), Encoding.UTF8))
                    {
                        worksheet.Save(sw);
                    }

                    contentTypes.Add(new XElement(ctNs + "Override",
                        new XAttribute("PartName", $"/xl/{sheetFileName}"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
                }

                contentTypes.Add(new XElement(ctNs + "Override",
                    new XAttribute("PartName", "/xl/workbook.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")));

                WriteXmlEntry(archive, "[Content_Types].xml", contentTypes);
                WriteXmlEntry(archive, "_rels/.rels", rels);
                WriteXmlEntry(archive, "xl/workbook.xml", workbook);
                WriteXmlEntry(archive, "xl/_rels/workbook.xml.rels", wbRels);
            }
        }

        private static void WriteXmlEntry(ZipArchive archive, string entryName, XElement element)
        {
            var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using (var sw = new StreamWriter(entry.Open(), Encoding.UTF8))
            {
                element.Save(sw);
            }
        }

        private static string GetColumnLetter(uint columnNumber)
        {
            string column = "";
            while (columnNumber > 0)
            {
                uint remainder = (columnNumber - 1) % 26;
                column = (char)('A' + remainder) + column;
                columnNumber = (columnNumber - 1) / 26;
            }
            return column;
        }

        private class PageResult
        {
            public int PageNumber { get; }
            public List<List<string>> Rows { get; }
            public PageResult(int pageNumber, List<List<string>> rows)
            {
                PageNumber = pageNumber;
                Rows = rows;
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
            }
            GC.SuppressFinalize(this);
        }
    }
}