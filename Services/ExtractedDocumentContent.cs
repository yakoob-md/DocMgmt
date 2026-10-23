namespace DocMgmt.Services;

/// <summary>Everything the preview partial needs to draw a .docx or .xlsx file.</summary>
public class ExtractedDocumentContent
{
	public bool IsExcel { get; set; }

	// Word content, in reading order.
	public List<WordBlock> Blocks { get; } = new();

	// Excel content, one entry per visible worksheet.
	public List<SheetPreview> Sheets { get; } = new();

	public bool IsTruncated { get; set; }

	public string? Notice { get; set; }
}

public abstract class WordBlock { }

public sealed class WordParagraph : WordBlock
{
	public int HeadingLevel { get; set; }   // 0 = normal paragraph, 1-6 = heading
	public bool IsListItem { get; set; }
	public int ListLevel { get; set; }
	public List<WordRun> Runs { get; } = new();
}

public sealed class WordRun
{
	public string Text { get; set; } = string.Empty;
	public bool Bold { get; set; }
	public bool Italic { get; set; }
	public bool Underline { get; set; }
}

public sealed class WordTable : WordBlock
{
	public List<List<WordTableCell>> Rows { get; } = new();
}

public sealed class WordTableCell
{
	public string Text { get; set; } = string.Empty;
	public int ColSpan { get; set; } = 1;
}

public sealed class SheetPreview
{
	public string Name { get; set; } = string.Empty;
	public List<SheetRow> Rows { get; } = new();
	public int ColumnCount { get; set; }
	public bool IsTruncated { get; set; }
}

public sealed class SheetRow
{
	public int Number { get; set; }
	public string[] Cells { get; set; } = Array.Empty<string>();
}