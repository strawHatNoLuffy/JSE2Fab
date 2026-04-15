using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WindowsFormTool.TskUtil
{
    public class TskDataFormatConversionProcessor : ITskProcessor
    {
        private const int MaxBatchCount = 25;
        private const int CommonColumnCount = 10;
        private const int DefaultTestValueStartIndex = 6;

        private static readonly string[] CommonHeaders =
        {
            "WAFER_ID",
            "LOT_ID",
            "DIE_X",
            "DIE_Y",
            "ULT",
            "SITE",
            "HBIN",
            "SBIN",
            "TIME",
            "TP_VERSION"
        };

        private static readonly Regex NumericValueRegex = new Regex(
            @"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?",
            RegexOptions.Compiled);

        private static readonly Regex ProgramSegmentRegex = new Regex(
            @"PROGRAM[\\/](?<name>[^\\/]+)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public void ProcessSingle(string firstFile, string secondFile, Action<string> updateStatus, ProgressBar progressBar = null)
        {
            if (string.IsNullOrWhiteSpace(firstFile))
            {
                MessageBox.Show(@"请先选择CSV文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ProcessBatch(new List<string> { firstFile }, null, updateStatus, progressBar);
        }

        public void ProcessBatch(List<string> firstFiles, List<string> secondFiles, Action<string> updateStatus, ProgressBar progressBar = null)
        {
            if (firstFiles == null || firstFiles.Count == 0)
            {
                MessageBox.Show(@"请先选择CSV文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (firstFiles.Count > MaxBatchCount)
            {
                MessageBox.Show(@"一次最多选择25个CSV文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            updateStatus = updateStatus ?? delegate { };

            var outputDirectory = TskFileHelper.GetSaveFilePath();
            var successCount = 0;
            var failureMessages = new List<string>();

            if (progressBar != null)
            {
                progressBar.Maximum = firstFiles.Count;
                progressBar.Value = 0;
            }

            foreach (var csvPath in firstFiles)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(csvPath) || !File.Exists(csvPath))
                    {
                        RecordFailure(updateStatus, failureMessages, csvPath, "文件不存在");
                        continue;
                    }

                    if (!string.Equals(Path.GetExtension(csvPath), ".csv", StringComparison.OrdinalIgnoreCase))
                    {
                        RecordFailure(updateStatus, failureMessages, csvPath, "不是CSV文件");
                        continue;
                    }

                    updateStatus(string.Format("开始转换：{0}\n", Path.GetFileName(csvPath)));

                    SourceCsvDocument document;
                    string errorMessage;
                    if (!TryParseSourceCsv(csvPath, out document, out errorMessage))
                    {
                        RecordFailure(updateStatus, failureMessages, csvPath, errorMessage);
                        continue;
                    }

                    string outputFileName;
                    if (!TryBuildOutputFileName(document, out outputFileName, out errorMessage))
                    {
                        RecordFailure(updateStatus, failureMessages, csvPath, errorMessage);
                        continue;
                    }

                    var outputPath = GetUniqueOutputPath(outputDirectory, outputFileName);
                    if (!string.Equals(outputPath, Path.Combine(outputDirectory, outputFileName), StringComparison.OrdinalIgnoreCase))
                    {
                        updateStatus(string.Format("输出文件已存在，自动重命名为：{0}\n", Path.GetFileName(outputPath)));
                    }

                    WriteOutputCsv(document, outputPath);
                    updateStatus(string.Format("已生成：{0}\n", outputPath));
                    successCount++;
                }
                catch (Exception ex)
                {
                    RecordFailure(updateStatus, failureMessages, csvPath, ex.Message);
                }
                finally
                {
                    if (progressBar != null && progressBar.Value < progressBar.Maximum)
                    {
                        progressBar.Value++;
                    }
                }
            }

            updateStatus(string.Format("转换完成：成功 {0} 个，失败 {1} 个。\n", successCount, failureMessages.Count));
            if (failureMessages.Count > 0)
            {
                updateStatus("失败明细：\n");
                foreach (var failureMessage in failureMessages)
                {
                    updateStatus(string.Format("- {0}\n", failureMessage));
                }
            }
        }

        private static void RecordFailure(Action<string> updateStatus, ICollection<string> failureMessages, string csvPath, string reason)
        {
            var fileName = string.IsNullOrWhiteSpace(csvPath) ? "<未命名文件>" : Path.GetFileName(csvPath);
            var message = string.Format("{0}：{1}", fileName, reason);
            failureMessages.Add(message);
            updateStatus(string.Format("转换失败：{0}\n", message));
        }

        private static bool TryParseSourceCsv(string csvPath, out SourceCsvDocument document, out string errorMessage)
        {
            document = null;
            errorMessage = null;

            var lines = File.ReadAllLines(csvPath);
            if (lines.Length == 0)
            {
                errorMessage = "文件为空";
                return false;
            }

            var metadata = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var layout = default(SourceLayout);
            List<string> testNameColumns = null;
            List<string> minLimitColumns = null;
            List<string> maxLimitColumns = null;
            var dataHeaderIndex = -1;

            for (var index = 0; index < lines.Length; index++)
            {
                var columns = SplitCsvLine(lines[index]);
                if (columns.Count == 0 || columns.All(string.IsNullOrWhiteSpace))
                {
                    continue;
                }

                if (TryDetectLayout(columns, out layout))
                {
                    dataHeaderIndex = index;
                    if (layout.UsesHeaderTestNames)
                    {
                        testNameColumns = columns;
                    }
                    break;
                }

                if (IsLegacyTestNameRow(columns))
                {
                    testNameColumns = columns;
                    continue;
                }

                if (IsLegacyMinLimitRow(columns))
                {
                    minLimitColumns = columns;
                    continue;
                }

                if (IsLegacyMaxLimitRow(columns))
                {
                    maxLimitColumns = columns;
                    continue;
                }

                ParseMetadataColumns(columns, metadata);
            }

            if (dataHeaderIndex < 0)
            {
                errorMessage = "缺少数据表头";
                return false;
            }

            var dataRows = new List<SourceDataRow>();
            for (var index = dataHeaderIndex + 1; index < lines.Length; index++)
            {
                var columns = SplitCsvLine(lines[index]);
                if (columns.Count == 0 || columns.All(string.IsNullOrWhiteSpace))
                {
                    continue;
                }

                var firstColumn = GetColumn(columns, 0);
                if (layout.UsesHeaderTestNames)
                {
                    if (IsMinRow(firstColumn))
                    {
                        minLimitColumns = columns;
                        continue;
                    }

                    if (IsMaxRow(firstColumn))
                    {
                        maxLimitColumns = columns;
                        continue;
                    }

                    if (IsBiasRow(firstColumn))
                    {
                        continue;
                    }
                }

                if (!IsDataRow(columns, layout))
                {
                    continue;
                }

                dataRows.Add(new SourceDataRow
                {
                    Site = GetColumn(columns, layout.SiteIndex),
                    Bin = GetColumn(columns, layout.BinIndex),
                    Sbin = layout.SbinIndex >= 0 ? GetColumn(columns, layout.SbinIndex) : GetColumn(columns, layout.BinIndex),
                    DieX = GetColumn(columns, layout.DieXIndex),
                    DieY = GetColumn(columns, layout.DieYIndex),
                    TestValues = GetAlignedColumns(columns, layout.TestValueStartIndex, GetTestCount(testNameColumns, minLimitColumns, maxLimitColumns, layout), false)
                });
            }

            if (testNameColumns == null)
            {
                errorMessage = "缺少测试项表头";
                return false;
            }

            if (minLimitColumns == null)
            {
                errorMessage = "缺少 MIN 行";
                return false;
            }

            if (maxLimitColumns == null)
            {
                errorMessage = "缺少 MAX 行";
                return false;
            }

            if (dataRows.Count == 0)
            {
                errorMessage = "未识别到数据行";
                return false;
            }

            var testCount = GetTestCount(testNameColumns, minLimitColumns, maxLimitColumns, layout);
            if (testCount <= 0)
            {
                errorMessage = "未识别到测试项列";
                return false;
            }

            var lotId = FirstNonEmpty(GetMetadataValue(metadata, "Lot"), GetMetadataValue(metadata, "Lot ID"));
            var waferId = FirstNonEmpty(GetMetadataValue(metadata, "Serial ID"), GetMetadataValue(metadata, "Wafer ID"));
            var testMachine = FirstNonEmpty(GetMetadataValue(metadata, "TestMachine"), GetMetadataValue(metadata, "Station"));
            var computerName = GetMetadataValue(metadata, "ComputerName");
            var proberCardId = GetMetadataValue(metadata, "Prober Card ID");
            var startTimeRaw = FirstNonEmpty(GetMetadataValue(metadata, "Start Time"), GetMetadataValue(metadata, "Beginning Time"));
            var tpVersionSource = FirstNonEmpty(GetMetadataValue(metadata, "TestFileName"), GetMetadataValue(metadata, "TST File Name For DTA"));

            if (string.IsNullOrWhiteSpace(lotId))
            {
                errorMessage = "缺少 Lot";
                return false;
            }

            if (string.IsNullOrWhiteSpace(waferId))
            {
                errorMessage = "缺少 Serial ID";
                return false;
            }

            if (string.IsNullOrWhiteSpace(testMachine))
            {
                errorMessage = "缺少 TestMachine";
                return false;
            }

            if (string.IsNullOrWhiteSpace(computerName))
            {
                errorMessage = "缺少 ComputerName";
                return false;
            }

            DateTime startTime;
            if (!TryParseDateTimeValue(startTimeRaw, out startTime))
            {
                errorMessage = "Start Time 格式无效";
                return false;
            }

            string tpVersion;
            if (!TryExtractTpVersion(tpVersionSource, out tpVersion))
            {
                errorMessage = "缺少 TP_VERSION 来源字段";
                return false;
            }

            document = new SourceCsvDocument
            {
                CsvPath = csvPath,
                LotId = lotId,
                WaferId = waferId,
                TestMachine = testMachine,
                ComputerName = computerName,
                ProberCardId = string.IsNullOrWhiteSpace(proberCardId) ? "NA" : proberCardId,
                TpVersion = tpVersion,
                StartTime = startTime,
                TestNames = GetAlignedColumns(testNameColumns, layout.TestValueStartIndex, testCount, false),
                MaxLimits = GetAlignedColumns(maxLimitColumns, layout.TestValueStartIndex, testCount, true),
                MinLimits = GetAlignedColumns(minLimitColumns, layout.TestValueStartIndex, testCount, true),
                DataRows = dataRows
            };

            return true;
        }

        private static bool TryBuildOutputFileName(SourceCsvDocument document, out string outputFileName, out string errorMessage)
        {
            outputFileName = null;
            errorMessage = null;

            if (string.IsNullOrWhiteSpace(document.TpVersion))
            {
                errorMessage = "缺少 TP_VERSION";
                return false;
            }

            if (string.IsNullOrWhiteSpace(document.TestMachine))
            {
                errorMessage = "缺少 TestMachine";
                return false;
            }

            if (string.IsNullOrWhiteSpace(document.ComputerName))
            {
                errorMessage = "缺少 ComputerName";
                return false;
            }

            if (string.IsNullOrWhiteSpace(document.WaferId))
            {
                errorMessage = "缺少 Serial ID";
                return false;
            }

            var timeCode = document.StartTime.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
            outputFileName = string.Format(
                CultureInfo.InvariantCulture,
                "IO#{0}#{1}#{2}#OI{3}-{4}_ALL_{5}.csv",
                document.TpVersion,
                document.TestMachine,
                string.IsNullOrWhiteSpace(document.ProberCardId) ? "NA" : document.ProberCardId,
                document.ComputerName,
                document.WaferId,
                timeCode);
            return true;
        }

        private static void WriteOutputCsv(SourceCsvDocument document, string outputPath)
        {
            using (var writer = new StreamWriter(outputPath, false, new UTF8Encoding(false)))
            {
                var headerRow = new List<string>(CommonHeaders);
                headerRow.AddRange(document.TestNames);
                WriteCsvLine(writer, headerRow);

                var limitHiRow = new List<string> { "#LIMIT_HI" };
                limitHiRow.AddRange(Enumerable.Repeat(string.Empty, CommonColumnCount - 1));
                limitHiRow.AddRange(document.MaxLimits);
                WriteCsvLine(writer, limitHiRow);

                var limitLoRow = new List<string> { "#LIMIT_LO" };
                limitLoRow.AddRange(Enumerable.Repeat(string.Empty, CommonColumnCount - 1));
                limitLoRow.AddRange(document.MinLimits);
                WriteCsvLine(writer, limitLoRow);

                foreach (var dataRow in document.DataRows)
                {
                    var outputRow = new List<string>
                    {
                        document.WaferId,
                        document.LotId,
                        dataRow.DieX,
                        dataRow.DieY,
                        BuildUlt(document.LotId, document.WaferId, dataRow.DieX, dataRow.DieY),
                        dataRow.Site,
                        dataRow.Bin,
                        dataRow.Sbin,
                        document.StartTime.ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture),
                        document.TpVersion
                    };

                    outputRow.AddRange(dataRow.TestValues);
                    TrimTrailingEmptyValues(outputRow, CommonColumnCount);
                    WriteCsvLine(writer, outputRow);
                }
            }
        }

        private static string BuildUlt(string lotId, string waferId, string dieX, string dieY)
        {
            return string.Format(CultureInfo.InvariantCulture, "{0}:{1}:{2}:{3}", lotId, waferId, dieX, dieY);
        }

        private static void TrimTrailingEmptyValues(List<string> values, int minimumCount)
        {
            while (values.Count > minimumCount && string.IsNullOrWhiteSpace(values[values.Count - 1]))
            {
                values.RemoveAt(values.Count - 1);
            }
        }

        private static void WriteCsvLine(TextWriter writer, IEnumerable<string> values)
        {
            writer.WriteLine(string.Join(",", values.Select(EscapeCsvValue).ToArray()));
        }

        private static string EscapeCsvValue(string value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            if (value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) < 0)
            {
                return value;
            }

            return string.Format(CultureInfo.InvariantCulture, "\"{0}\"", value.Replace("\"", "\"\""));
        }

        private static List<string> SplitCsvLine(string line)
        {
            var values = new List<string>();
            if (line == null)
            {
                return values;
            }

            var builder = new StringBuilder();
            var inQuotes = false;

            for (var index = 0; index < line.Length; index++)
            {
                var current = line[index];
                if (current == '"')
                {
                    if (inQuotes && index + 1 < line.Length && line[index + 1] == '"')
                    {
                        builder.Append('"');
                        index++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }

                    continue;
                }

                if (current == ',' && !inQuotes)
                {
                    values.Add(builder.ToString().Trim());
                    builder.Clear();
                    continue;
                }

                builder.Append(current);
            }

            values.Add(builder.ToString().Trim());
            return values;
        }

        private static bool TryDetectLayout(IReadOnlyList<string> columns, out SourceLayout layout)
        {
            if (string.Equals(GetColumn(columns, 0), "Site", StringComparison.OrdinalIgnoreCase)
                && string.Equals(GetColumn(columns, 1), "Serial", StringComparison.OrdinalIgnoreCase)
                && string.Equals(GetColumn(columns, 2), "Sbin", StringComparison.OrdinalIgnoreCase)
                && string.Equals(GetColumn(columns, 3), "Bin", StringComparison.OrdinalIgnoreCase)
                && string.Equals(GetColumn(columns, 4), "X", StringComparison.OrdinalIgnoreCase)
                && string.Equals(GetColumn(columns, 5), "Y", StringComparison.OrdinalIgnoreCase))
            {
                layout = new SourceLayout(false, 0, 2, 3, 4, 5, DefaultTestValueStartIndex, 0);
                return true;
            }

            if (string.Equals(GetColumn(columns, 0), "TestNo", StringComparison.OrdinalIgnoreCase)
                && string.Equals(GetColumn(columns, 1), "SiteNo", StringComparison.OrdinalIgnoreCase)
                && string.Equals(GetColumn(columns, 2), "Bin", StringComparison.OrdinalIgnoreCase)
                && string.Equals(GetColumn(columns, 3), "Time/mS", StringComparison.OrdinalIgnoreCase)
                && string.Equals(GetColumn(columns, 4), "X", StringComparison.OrdinalIgnoreCase)
                && string.Equals(GetColumn(columns, 5), "Y", StringComparison.OrdinalIgnoreCase))
            {
                layout = new SourceLayout(true, 1, -1, 2, 4, 5, DefaultTestValueStartIndex, 0);
                return true;
            }

            layout = default(SourceLayout);
            return false;
        }

        private static bool IsLegacyTestNameRow(IReadOnlyList<string> columns)
        {
            return string.Equals(GetColumn(columns, 0), "Test Name", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsLegacyMinLimitRow(IReadOnlyList<string> columns)
        {
            return string.Equals(GetColumn(columns, 0), "Min Limit", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsLegacyMaxLimitRow(IReadOnlyList<string> columns)
        {
            return string.Equals(GetColumn(columns, 0), "Max Limit", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsMinRow(string firstColumn)
        {
            return string.Equals(firstColumn, "MIN", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsMaxRow(string firstColumn)
        {
            return string.Equals(firstColumn, "MAX", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsBiasRow(string firstColumn)
        {
            return firstColumn.StartsWith("BIAS", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsDataRow(IReadOnlyList<string> columns, SourceLayout layout)
        {
            int numericMarker;
            return columns.Count >= layout.TestValueStartIndex
                && int.TryParse(GetColumn(columns, layout.RowMarkerIndex), NumberStyles.Integer, CultureInfo.InvariantCulture, out numericMarker);
        }

        private static void ParseMetadataColumns(IReadOnlyList<string> columns, IDictionary<string, string> metadata)
        {
            var foundExplicitKey = false;
            for (var index = 0; index < columns.Count; index++)
            {
                var cell = GetColumn(columns, index);
                if (!LooksLikeMetadataKey(cell))
                {
                    continue;
                }

                foundExplicitKey = true;
                var key = NormalizeMetadataKey(cell);
                var value = GetFirstMetadataValue(columns, index + 1);
                if (!metadata.ContainsKey(key))
                {
                    metadata[key] = value;
                }
            }

            if (foundExplicitKey)
            {
                return;
            }

            var fallbackKey = NormalizeMetadataKey(GetColumn(columns, 0));
            if (string.IsNullOrWhiteSpace(fallbackKey) || metadata.ContainsKey(fallbackKey))
            {
                return;
            }

            metadata[fallbackKey] = GetFirstMetadataValue(columns, 1);
        }

        private static bool LooksLikeMetadataKey(string value)
        {
            return !string.IsNullOrWhiteSpace(value) && value.Trim().EndsWith(":", StringComparison.Ordinal);
        }

        private static string NormalizeMetadataKey(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim().TrimEnd(':').Trim();
        }

        private static string GetFirstMetadataValue(IReadOnlyList<string> columns, int startIndex)
        {
            for (var index = startIndex; index < columns.Count; index++)
            {
                var candidate = GetColumn(columns, index);
                if (string.IsNullOrWhiteSpace(candidate))
                {
                    continue;
                }

                if (LooksLikeMetadataKey(candidate))
                {
                    break;
                }

                return candidate;
            }

            return string.Empty;
        }

        private static string GetMetadataValue(IDictionary<string, string> metadata, string key)
        {
            string value;
            return metadata.TryGetValue(key, out value) ? value : string.Empty;
        }

        private static string FirstNonEmpty(params string[] values)
        {
            return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? string.Empty;
        }

        private static int GetTestCount(IReadOnlyList<string> testNameColumns, IReadOnlyList<string> minLimitColumns, IReadOnlyList<string> maxLimitColumns, SourceLayout layout)
        {
            return Math.Max(
                0,
                Math.Max(
                    testNameColumns == null ? 0 : testNameColumns.Count,
                    Math.Max(minLimitColumns == null ? 0 : minLimitColumns.Count, maxLimitColumns == null ? 0 : maxLimitColumns.Count))
                - layout.TestValueStartIndex);
        }

        private static List<string> GetAlignedColumns(IReadOnlyList<string> columns, int startIndex, int itemCount, bool normalizeNumericValues)
        {
            var values = new List<string>();
            for (var index = 0; index < itemCount; index++)
            {
                var value = GetColumn(columns, startIndex + index);
                values.Add(normalizeNumericValues ? NormalizeLimitValue(value) : value);
            }

            return values;
        }

        private static string NormalizeLimitValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var match = NumericValueRegex.Match(value);
            return match.Success ? match.Value : value;
        }

        private static string GetColumn(IReadOnlyList<string> columns, int index)
        {
            if (columns == null || index < 0 || index >= columns.Count)
            {
                return string.Empty;
            }

            return columns[index] == null ? string.Empty : columns[index].Trim();
        }

        private static bool TryExtractTpVersion(string sourceValue, out string tpVersion)
        {
            tpVersion = string.Empty;
            if (string.IsNullOrWhiteSpace(sourceValue))
            {
                return false;
            }

            var match = ProgramSegmentRegex.Match(sourceValue);
            if (match.Success)
            {
                tpVersion = match.Groups["name"].Value.Trim();
                return !string.IsNullOrWhiteSpace(tpVersion);
            }

            tpVersion = RemoveExtension(sourceValue);
            return !string.IsNullOrWhiteSpace(tpVersion);
        }

        private static string RemoveExtension(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return string.Empty;
            }

            return Path.GetFileNameWithoutExtension(fileName.Trim());
        }

        private static string GetUniqueOutputPath(string outputDirectory, string outputFileName)
        {
            var baseName = Path.GetFileNameWithoutExtension(outputFileName);
            var extension = Path.GetExtension(outputFileName);
            var candidatePath = Path.Combine(outputDirectory, outputFileName);

            if (!File.Exists(candidatePath))
            {
                return candidatePath;
            }

            var sequence = 1;
            do
            {
                candidatePath = Path.Combine(
                    outputDirectory,
                    string.Format(CultureInfo.InvariantCulture, "{0}_{1}{2}", baseName, sequence, extension));
                sequence++;
            }
            while (File.Exists(candidatePath));

            return candidatePath;
        }

        private static bool TryParseDateTimeValue(string value, out DateTime dateTime)
        {
            var formats = new[]
            {
                "yyyy/M/d H:mm",
                "yyyy/M/d H:mm:ss",
                "yyyy/M/d HH:mm",
                "yyyy/M/d HH:mm:ss",
                "yyyy-MM-dd HH:mm:ss",
                "yyyy-MM-dd H:mm"
            };

            if (DateTime.TryParseExact(value, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime))
            {
                return true;
            }

            return DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime)
                || DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, out dateTime);
        }

        private sealed class SourceCsvDocument
        {
            public string CsvPath { get; set; }

            public string LotId { get; set; }

            public string WaferId { get; set; }

            public string TestMachine { get; set; }

            public string ComputerName { get; set; }

            public string ProberCardId { get; set; }

            public string TpVersion { get; set; }

            public DateTime StartTime { get; set; }

            public List<string> TestNames { get; set; }

            public List<string> MaxLimits { get; set; }

            public List<string> MinLimits { get; set; }

            public List<SourceDataRow> DataRows { get; set; }
        }

        private sealed class SourceDataRow
        {
            public string Site { get; set; }

            public string Sbin { get; set; }

            public string Bin { get; set; }

            public string DieX { get; set; }

            public string DieY { get; set; }

            public List<string> TestValues { get; set; }
        }

        private struct SourceLayout
        {
            public SourceLayout(bool usesHeaderTestNames, int siteIndex, int sbinIndex, int binIndex, int dieXIndex, int dieYIndex, int testValueStartIndex, int rowMarkerIndex)
            {
                UsesHeaderTestNames = usesHeaderTestNames;
                SiteIndex = siteIndex;
                SbinIndex = sbinIndex;
                BinIndex = binIndex;
                DieXIndex = dieXIndex;
                DieYIndex = dieYIndex;
                TestValueStartIndex = testValueStartIndex;
                RowMarkerIndex = rowMarkerIndex;
            }

            public bool UsesHeaderTestNames { get; }

            public int SiteIndex { get; }

            public int SbinIndex { get; }

            public int BinIndex { get; }

            public int DieXIndex { get; }

            public int DieYIndex { get; }

            public int TestValueStartIndex { get; }

            public int RowMarkerIndex { get; }
        }
    }
}
