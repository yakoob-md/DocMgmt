using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace DocMgmt.Services;

/// <summary>
/// Reads .docx and .xlsx files using only the .NET base class library (System.IO.Compression + System.Xml).
/// Output is capped so a huge or hostile file cannot exhaust server memory.
/// </summary>
public class NativeDocumentParser
{
	private const long MaxPartBytes = 64L * 1024 * 1024; // max uncompressed size of any single XML part
	private const int MaxWordBlocks = 3000;
	private const int MaxTableRows = 500;
	private const int MaxSheets = 20;
	private const int MaxRowsPerSheet = 500;
	private const int MaxColumnsPerSheet = 50;

	// XXE-safe: DTDs are rejected and no external resolver is used.
	private static readonly XmlReaderSettings SafeXml = new()
	{
		DtdProcessing = DtdProcessing.Prohibit,
		XmlResolver = null,
		IgnoreComments = true,
		IgnoreProcessingInstructions = true
	};

	// ------------------------------------------------------------------ public API

	public ExtractedDocumentContent ParseFromStream(Stream stream, string extension)
	{
		if (stream.CanSeek) stream.Position = 0;

		using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

		return extension.ToLowerInvariant() switch
		{
			".xlsx" => ParseWorkbook(archive),
			".docx" => ParseWordDocument(archive),
			_ => throw new NotSupportedException($"Preview is not supported for '{extension}'.")
		};
	}

	/// <summary>True when the stream is a zip package that contains the main part of a Word/Excel file.</summary>
	public static bool HasRequiredParts(Stream stream, string extension)
	{
		try
		{
			if (stream.CanSeek) stream.Position = 0;
			using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
			var fallback = extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
				? "xl/workbook.xml"
				: "word/document.xml";
			return FindMainPart(archive, fallback).Entry is not null;
		}
		catch (Exception ex) when (ex is InvalidDataException or XmlException)
		{
			return false;
		}
		finally
		{
			if (stream.CanSeek) stream.Position = 0;
		}
	}

	// ------------------------------------------------------------------ WORD

	private ExtractedDocumentContent ParseWordDocument(ZipArchive archive)
	{
		var result = new ExtractedDocumentContent { IsExcel = false };

		var (mainPath, mainEntry) = FindMainPart(archive, "word/document.xml");
		if (mainEntry is null)
			throw new InvalidDataException("This is not a valid Word document (the main document part is missing).");

		var styles = ReadParagraphStyles(archive, mainPath);

		// PreserveWhitespace keeps runs such as <w:t xml:space="preserve"> </w:t>.
		var doc = LoadXml(mainEntry, LoadOptions.PreserveWhitespace);
		var body = doc.Root?.Kid("body");
		if (body is not null)
			CollectBlocks(body.Elements(), styles, result);

		if (result.IsTruncated)
			result.Notice = "This is a partial preview. Open the original file to see the full document.";

		return result;
	}

	private static void CollectBlocks(IEnumerable<XElement> elements, ParagraphStyles styles, ExtractedDocumentContent result)
	{
		foreach (var el in elements)
		{
			if (result.Blocks.Count >= MaxWordBlocks)
			{
				result.IsTruncated = true;
				return;
			}

			switch (el.Name.LocalName)
			{
				case "p":
					result.Blocks.Add(ReadParagraph(el, styles));
					break;
				case "tbl":
					result.Blocks.Add(ReadTable(el, result));
					break;
				case "sdt": // content controls wrap normal blocks
					var content = el.Kid("sdtContent");
					if (content is not null)
						CollectBlocks(content.Elements(), styles, result);
					break;
			}
		}
	}

	private static WordParagraph ReadParagraph(XElement p, ParagraphStyles styles)
	{
		var paragraph = new WordParagraph();
		var pPr = p.Kid("pPr");

		var styleId = pPr?.Kid("pStyle")?.Attr("val");
		if (styleId is not null && styles.HeadingLevels.TryGetValue(styleId, out var level))
			paragraph.HeadingLevel = level;

		if (int.TryParse(pPr?.Kid("outlineLvl")?.Attr("val"), out var outline) && outline is >= 0 and <= 8)
			paragraph.HeadingLevel = Math.Min(outline + 1, 6);

		// A paragraph is a list item when it (or its style) carries numbering. numId 0 means "no list".
		var numPr = pPr?.Kid("numPr");
		bool styleIsList = styleId is not null && styles.ListStyles.Contains(styleId);
		bool directList = numPr is not null && numPr.Kid("numId")?.Attr("val") != "0";
		bool directNoList = numPr is not null && numPr.Kid("numId")?.Attr("val") == "0";

		if (directList || (styleIsList && !directNoList))
		{
			paragraph.IsListItem = true;
			int.TryParse(numPr?.Kid("ilvl")?.Attr("val"), out var ilvl);
			paragraph.ListLevel = Math.Clamp(ilvl, 0, 8);
		}

		paragraph.Runs.AddRange(ReadRuns(p));
		return paragraph;
	}

	private static List<WordRun> ReadRuns(XElement p)
	{
		var runs = new List<WordRun>();

		foreach (var r in p.Descendants().Where(x => x.Name.LocalName == "r"))
		{
			// Skip runs that belong to a nested paragraph (e.g. text boxes) and mc:Fallback duplicates.
			if (!ReferenceEquals(r.Ancestors().FirstOrDefault(a => a.Name.LocalName == "p"), p)) continue;
			if (r.Ancestors().Any(a => a.Name.LocalName is "Fallback" or "del")) continue;

			var sb = new StringBuilder();
			foreach (var child in r.Elements())
			{
				switch (child.Name.LocalName)
				{
					case "t": sb.Append(child.Value); break;
					case "tab": sb.Append('\t'); break;
					case "br":
					case "cr": sb.Append('\n'); break;
					case "noBreakHyphen": sb.Append('-'); break;
				}
			}

			if (sb.Length == 0) continue;

			var rPr = r.Kid("rPr");
			bool bold = IsOn(rPr?.Kid("b"));
			bool italic = IsOn(rPr?.Kid("i"));
			var u = rPr?.Kid("u");
			bool underline = u is not null && !string.Equals(u.Attr("val"), "none", StringComparison.OrdinalIgnoreCase);

			var last = runs.Count > 0 ? runs[^1] : null;
			if (last is not null && last.Bold == bold && last.Italic == italic && last.Underline == underline)
				last.Text += sb.ToString();
			else
				runs.Add(new WordRun { Text = sb.ToString(), Bold = bold, Italic = italic, Underline = underline });
		}

		return runs;
	}

	private static WordTable ReadTable(XElement tbl, ExtractedDocumentContent result)
	{
		var table = new WordTable();

		foreach (var tr in tbl.Kids("tr"))
		{
			if (table.Rows.Count >= MaxTableRows)
			{
				result.IsTruncated = true;
				break;
			}

			var row = new List<WordTableCell>();
			foreach (var tc in tr.Kids("tc"))
			{
				var lines = tc.Descendants()
					.Where(x => x.Name.LocalName == "p" && !x.Ancestors().Any(a => a.Name.LocalName == "Fallback"))
					.Select(pp => string.Concat(ReadRuns(pp).Select(run => run.Text)));

				int.TryParse(tc.Kid("tcPr")?.Kid("gridSpan")?.Attr("val"), out var span);

				row.Add(new WordTableCell
				{
					Text = string.Join("\n", lines).Trim('\n'),
					ColSpan = Math.Clamp(span, 1, 50)
				});
			}
			table.Rows.Add(row);
		}

		return table;
	}

	private sealed class ParagraphStyles
	{
		public Dictionary<string, int> HeadingLevels { get; } = new(StringComparer.Ordinal);
		public HashSet<string> ListStyles { get; } = new(StringComparer.Ordinal);
	}

	private static ParagraphStyles ReadParagraphStyles(ZipArchive archive, string mainPath)
	{
		var map = new ParagraphStyles();

		var rels = ReadRelationships(archive, RelsPathFor(mainPath));
		var stylesRel = rels.Values.FirstOrDefault(r => r.Type.EndsWith("/styles", StringComparison.OrdinalIgnoreCase) && !r.External);
		var stylesPath = stylesRel is not null ? ResolvePath(PartDirectory(mainPath), stylesRel.Target) : "word/styles.xml";

		var entry = FindEntry(archive, stylesPath);
		if (entry is null) return map;

		var root = LoadXml(entry).Root;
		if (root is null) return map;

		foreach (var style in root.Kids("style"))
		{
			if (style.Attr("type") != "paragraph") continue;
			var id = style.Attr("styleId");
			if (id is null) continue;

			var name = style.Kid("name")?.Attr("val")?.ToLowerInvariant() ?? "";
			int level = 0;

			if (name == "title")
				level = 1;
			else if (name.StartsWith("heading ", StringComparison.Ordinal) && int.TryParse(name.AsSpan(8), out var n) && n is >= 1 and <= 9)
				level = n;
			else if (int.TryParse(style.Kid("pPr")?.Kid("outlineLvl")?.Attr("val"), out var ol) && ol is >= 0 and <= 8)
				level = ol + 1;

			if (level > 0) map.HeadingLevels[id] = Math.Min(level, 6);

			var styleNum = style.Kid("pPr")?.Kid("numPr");
			if (styleNum is not null && styleNum.Kid("numId")?.Attr("val") != "0")
				map.ListStyles.Add(id);
		}

		return map;
	}

	private static bool IsOn(XElement? toggle)
	{
		if (toggle is null) return false;
		var val = toggle.Attr("val");
		return !(string.Equals(val, "0", StringComparison.Ordinal) ||
				 string.Equals(val, "false", StringComparison.OrdinalIgnoreCase) ||
				 string.Equals(val, "off", StringComparison.OrdinalIgnoreCase));
	}

	// ------------------------------------------------------------------ EXCEL

	private ExtractedDocumentContent ParseWorkbook(ZipArchive archive)
	{
		var result = new ExtractedDocumentContent { IsExcel = true };

		var (wbPath, wbEntry) = FindMainPart(archive, "xl/workbook.xml");
		if (wbEntry is null)
			throw new InvalidDataException("This is not a valid Excel workbook (the workbook part is missing).");

		var wbDir = PartDirectory(wbPath);
		var rels = ReadRelationships(archive, RelsPathFor(wbPath));
		var wb = LoadXml(wbEntry).Root ?? throw new InvalidDataException("The workbook part is empty.");

		var date1904 = wb.Kid("workbookPr")?.Attr("date1904") is "1" or "true";

		string RelTarget(string typeSuffix, string fallback)
		{
			var rel = rels.Values.FirstOrDefault(r => r.Type.EndsWith(typeSuffix, StringComparison.OrdinalIgnoreCase) && !r.External);
			return rel is not null ? ResolvePath(wbDir, rel.Target) : fallback;
		}

		var sharedStrings = ReadSharedStrings(archive, RelTarget("/sharedStrings", "xl/sharedStrings.xml"));
		var dateStyles = ReadDateStyles(archive, RelTarget("/styles", "xl/styles.xml"));

		// Collect worksheets in workbook order, skipping hidden sheets and chart sheets.
		var sheets = new List<(string Name, ZipArchiveEntry Entry)>();
		foreach (var sheet in wb.Kid("sheets")?.Kids("sheet") ?? Enumerable.Empty<XElement>())
		{
			var state = sheet.Attr("state");
			if (state is "hidden" or "veryHidden") continue;

			var rid = sheet.Attr("id");
			if (rid is null || !rels.TryGetValue(rid, out var rel)) continue;
			if (!rel.Type.EndsWith("/worksheet", StringComparison.OrdinalIgnoreCase) || rel.External) continue;

			var entry = FindEntry(archive, ResolvePath(wbDir, rel.Target));
			if (entry is null) continue;

			sheets.Add((sheet.Attr("name") ?? $"Sheet{sheets.Count + 1}", entry));
		}

		// Fallback for workbooks whose relationships are missing or broken.
		if (sheets.Count == 0)
		{
			var pattern = new Regex(@"^xl/worksheets/sheet\d+\.xml$", RegexOptions.IgnoreCase);
			foreach (var e in archive.Entries.Where(e => pattern.IsMatch(e.FullName)).OrderBy(e => e.FullName, StringComparer.OrdinalIgnoreCase))
				sheets.Add((Path.GetFileNameWithoutExtension(e.Name), e));
		}

		if (sheets.Count > MaxSheets)
		{
			result.IsTruncated = true;
			sheets = sheets.Take(MaxSheets).ToList();
		}

		foreach (var (name, entry) in sheets)
		{
			var preview = ReadSheet(entry, name, sharedStrings, dateStyles, date1904);
			result.Sheets.Add(preview);
			if (preview.IsTruncated) result.IsTruncated = true;
		}

		if (result.IsTruncated)
			result.Notice = $"Preview limit: first {MaxRowsPerSheet} non-empty rows, {MaxColumnsPerSheet} columns and {MaxSheets} sheets. Open the original file to see everything.";

		return result;
	}

	private static List<string> ReadSharedStrings(ZipArchive archive, string path)
	{
		var list = new List<string>();
		var entry = FindEntry(archive, path);
		if (entry is null) return list;

		using var stream = OpenPart(entry);
		using var reader = XmlReader.Create(stream, SafeXml);
		reader.MoveToContent();

		while (!reader.EOF)
		{
			if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si")
				list.Add(ReadRichText((XElement)XNode.ReadFrom(reader)));
			else
				reader.Read();
		}

		return list;
	}

	// A string item is either <t> or several <r><t> runs. <rPh> (phonetic hints) must be ignored.
	private static string ReadRichText(XElement item)
	{
		var sb = new StringBuilder();
		foreach (var child in item.Elements())
		{
			if (child.Name.LocalName == "t")
				sb.Append(child.Value);
			else if (child.Name.LocalName == "r")
				foreach (var t in child.Kids("t")) sb.Append(t.Value);
		}
		return sb.ToString();
	}

	private static bool[] ReadDateStyles(ZipArchive archive, string path)
	{
		var entry = FindEntry(archive, path);
		if (entry is null) return Array.Empty<bool>();

		var root = LoadXml(entry).Root;
		if (root is null) return Array.Empty<bool>();

		var custom = new Dictionary<int, string>();
		foreach (var nf in root.Kid("numFmts")?.Kids("numFmt") ?? Enumerable.Empty<XElement>())
			if (int.TryParse(nf.Attr("numFmtId"), out var id))
				custom[id] = nf.Attr("formatCode") ?? "";

		var flags = new List<bool>();
		foreach (var xf in root.Kid("cellXfs")?.Kids("xf") ?? Enumerable.Empty<XElement>())
		{
			int.TryParse(xf.Attr("numFmtId"), out var fmtId);
			flags.Add(IsDateFormat(fmtId, custom));
		}
		return flags.ToArray();
	}

	private static bool IsDateFormat(int id, Dictionary<int, string> custom)
	{
		if (id is >= 14 and <= 22 or >= 27 and <= 36 or >= 45 and <= 47 or >= 50 and <= 58) return true;
		return custom.TryGetValue(id, out var code) && LooksLikeDateFormat(code);
	}

	private static bool LooksLikeDateFormat(string code)
	{
		bool inQuote = false, inBracket = false;
		for (int i = 0; i < code.Length; i++)
		{
			char c = code[i];
			if (inQuote) { if (c == '"') inQuote = false; continue; }
			if (inBracket) { if (c == ']') inBracket = false; continue; }

			switch (c)
			{
				case '"': inQuote = true; break;
				case '[': inBracket = true; break;
				case '\\':
				case '_':
				case '*': i++; break; // next char is a literal / padding
				case 'y':
				case 'Y':
				case 'm':
				case 'M':
				case 'd':
				case 'D':
				case 'h':
				case 'H':
				case 's':
				case 'S':
					return true;
			}
		}
		return false;
	}

	private static SheetPreview ReadSheet(ZipArchiveEntry entry, string name, List<string> sharedStrings, bool[] dateStyles, bool date1904)
	{
		var sheet = new SheetPreview { Name = name };
		int maxCol = 0;
		int lastRow = 0;

		using var stream = OpenPart(entry);
		using var reader = XmlReader.Create(stream, SafeXml);
		reader.MoveToContent();

		while (!reader.EOF)
		{
			if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row")
			{
				reader.Read();
				continue;
			}

			var rowEl = (XElement)XNode.ReadFrom(reader); // reader is now past this row

			int rowNumber = int.TryParse(rowEl.Attr("r"), out var rn) ? rn : lastRow + 1;
			lastRow = rowNumber;

			var cells = new List<string>();
			int nextCol = 0;
			bool hasValue = false;

			foreach (var c in rowEl.Kids("c"))
			{
				int col = ParseColumnIndex(c.Attr("r"));
				if (col < 0) col = nextCol;
				nextCol = col + 1;

				if (col >= MaxColumnsPerSheet)
				{
					sheet.IsTruncated = true;
					continue;
				}

				var text = ReadCellText(c, sharedStrings, dateStyles, date1904);
				if (text.Length == 0) continue;

				while (cells.Count <= col) cells.Add(string.Empty); // keep blank cells so columns stay aligned
				cells[col] = text;
				hasValue = true;
			}

			if (!hasValue) continue; // ignore empty, formatting-only rows

			if (sheet.Rows.Count >= MaxRowsPerSheet)
			{
				sheet.IsTruncated = true;
				break;
			}

			maxCol = Math.Max(maxCol, cells.Count);
			sheet.Rows.Add(new SheetRow { Number = rowNumber, Cells = cells.ToArray() });
		}

		// Pad every row to the same width.
		sheet.ColumnCount = maxCol;
		foreach (var row in sheet.Rows)
		{
			if (row.Cells.Length == maxCol) continue;
			var padded = new string[maxCol];
			Array.Fill(padded, string.Empty);
			Array.Copy(row.Cells, padded, row.Cells.Length);
			row.Cells = padded;
		}

		return sheet;
	}

	private static string ReadCellText(XElement c, List<string> sharedStrings, bool[] dateStyles, bool date1904)
	{
		var type = c.Attr("t");
		var v = c.Kid("v")?.Value;

		switch (type)
		{
			case "s":
				return int.TryParse(v, out var idx) && idx >= 0 && idx < sharedStrings.Count ? sharedStrings[idx] : string.Empty;
			case "inlineStr":
				var inline = c.Kid("is");
				return inline is null ? string.Empty : ReadRichText(inline);
			case "b":
				return v is null ? string.Empty : (v == "1" ? "TRUE" : "FALSE");
			case "str":
			case "e":
			case "d":
				return v ?? string.Empty;
		}

		// Numbers (and anything untyped).
		if (string.IsNullOrWhiteSpace(v)) return string.Empty;
		if (!double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || !double.IsFinite(number))
			return v;

		if (int.TryParse(c.Attr("s"), out var style) && style >= 0 && style < dateStyles.Length && dateStyles[style])
		{
			var date = FormatExcelDate(number, date1904);
			if (date is not null) return date;
		}

		if (Math.Abs(number) >= 1e15 || (number != 0 && Math.Abs(number) < 1e-9)) return v;
		return number.ToString("G15", CultureInfo.InvariantCulture); // hides 0.30000000000000004 style noise
	}

	private static string? FormatExcelDate(double serial, bool date1904)
	{
		if (serial < 0 || serial >= 2958466) return null;

		double days = Math.Floor(serial);
		double seconds = Math.Round((serial - days) * 86400);

		DateTime origin = date1904 ? new DateTime(1904, 1, 1)
			: serial < 60 ? new DateTime(1899, 12, 31)   // Excel treats 1900 as a leap year; serials >= 60 are shifted by one day
			: new DateTime(1899, 12, 30);

		var value = origin.AddDays(days).AddSeconds(seconds);

		if (seconds == 0) return value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
		if (serial < 1) return value.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
		return value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
	}

	private static int ParseColumnIndex(string? cellRef)
	{
		if (string.IsNullOrEmpty(cellRef)) return -1;

		int col = 0, letters = 0;
		foreach (char ch in cellRef)
		{
			if (ch is >= 'A' and <= 'Z') col = col * 26 + (ch - 'A' + 1);
			else if (ch is >= 'a' and <= 'z') col = col * 26 + (ch - 'a' + 1);
			else break;
			letters++;
			if (letters > 3) return -1;
		}
		return letters == 0 ? -1 : col - 1;
	}

	// ------------------------------------------------------------------ OPC package helpers

	private sealed record Relationship(string Id, string Type, string Target, bool External);

	private static (string Path, ZipArchiveEntry? Entry) FindMainPart(ZipArchive archive, string fallbackPath)
	{
		var rels = ReadRelationships(archive, "_rels/.rels");
		var main = rels.Values.FirstOrDefault(r => r.Type.EndsWith("/officeDocument", StringComparison.OrdinalIgnoreCase) && !r.External);

		if (main is not null)
		{
			var path = ResolvePath(string.Empty, main.Target);
			var entry = FindEntry(archive, path);
			if (entry is not null) return (path, entry);
		}

		return (fallbackPath, FindEntry(archive, fallbackPath));
	}

	private static Dictionary<string, Relationship> ReadRelationships(ZipArchive archive, string relsPath)
	{
		var map = new Dictionary<string, Relationship>(StringComparer.Ordinal);
		var entry = FindEntry(archive, relsPath);
		if (entry is null) return map;

		var root = LoadXml(entry).Root;
		if (root is null) return map;

		foreach (var rel in root.Kids("Relationship"))
		{
			var id = rel.Attr("Id");
			var type = rel.Attr("Type");
			var target = rel.Attr("Target");
			if (id is null || type is null || target is null) continue;

			map[id] = new Relationship(id, type, target,
				string.Equals(rel.Attr("TargetMode"), "External", StringComparison.OrdinalIgnoreCase));
		}
		return map;
	}

	private static string PartDirectory(string path)
	{
		int i = path.LastIndexOf('/');
		return i < 0 ? string.Empty : path[..i];
	}

	private static string RelsPathFor(string partPath)
	{
		var dir = PartDirectory(partPath);
		var file = partPath[(dir.Length == 0 ? 0 : dir.Length + 1)..];
		return (dir.Length == 0 ? string.Empty : dir + "/") + "_rels/" + file + ".rels";
	}

	private static string ResolvePath(string baseDir, string target)
	{
		target = target.Replace('\\', '/');
		var combined = target.StartsWith('/') ? target : (baseDir.Length == 0 ? target : baseDir + "/" + target);

		var parts = new List<string>();
		foreach (var seg in combined.Split('/', StringSplitOptions.RemoveEmptyEntries))
		{
			if (seg == ".") continue;
			if (seg == "..") { if (parts.Count > 0) parts.RemoveAt(parts.Count - 1); continue; }
			parts.Add(seg);
		}
		return string.Join('/', parts);
	}

	private static ZipArchiveEntry? FindEntry(ZipArchive archive, string path)
	{
		path = path.TrimStart('/').Replace('\\', '/');
		return archive.GetEntry(path) ??
			   archive.Entries.FirstOrDefault(e => string.Equals(e.FullName.Replace('\\', '/'), path, StringComparison.OrdinalIgnoreCase));
	}

	private static Stream OpenPart(ZipArchiveEntry entry)
	{
		if (entry.Length > MaxPartBytes)
			throw new InvalidDataException($"The part '{entry.FullName}' is too large to preview.");
		return new LimitedReadStream(entry.Open(), MaxPartBytes);
	}

	private static XDocument LoadXml(ZipArchiveEntry entry, LoadOptions options = LoadOptions.None)
	{
		using var stream = OpenPart(entry);
		using var reader = XmlReader.Create(stream, SafeXml);
		return XDocument.Load(reader, options);
	}

	/// <summary>Stops reading after a fixed number of bytes, even if the zip header lied about the size (zip bombs).</summary>
	private sealed class LimitedReadStream : Stream
	{
		private readonly Stream _inner;
		private readonly long _limit;
		private long _total;

		public LimitedReadStream(Stream inner, long limit) { _inner = inner; _limit = limit; }

		public override int Read(byte[] buffer, int offset, int count) => Track(_inner.Read(buffer, offset, count));
		public override int Read(Span<byte> buffer) => Track(_inner.Read(buffer));

		private int Track(int read)
		{
			_total += read;
			if (_total > _limit) throw new InvalidDataException("The file part is larger than the preview limit.");
			return read;
		}

		public override bool CanRead => true;
		public override bool CanSeek => false;
		public override bool CanWrite => false;
		public override long Length => throw new NotSupportedException();
		public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
		public override void Flush() { }
		public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
		public override void SetLength(long value) => throw new NotSupportedException();
		public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

		protected override void Dispose(bool disposing)
		{
			if (disposing) _inner.Dispose();
			base.Dispose(disposing);
		}
	}
}

internal static class XmlLocalNameExtensions
{
	// Namespace-agnostic helpers so both "transitional" and "strict" OOXML parse the same way.
	public static IEnumerable<XElement> Kids(this XElement e, string localName) =>
		e.Elements().Where(x => x.Name.LocalName == localName);

	public static XElement? Kid(this XElement e, string localName) =>
		e.Elements().FirstOrDefault(x => x.Name.LocalName == localName);

	public static string? Attr(this XElement e, string localName) =>
		e.Attributes().FirstOrDefault(a => a.Name.LocalName == localName)?.Value;
}