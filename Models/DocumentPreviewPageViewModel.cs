using DocMgmt.Services;

namespace DocMgmt.Models;

/// <summary>View model for the full-page ("open in new tab") preview of a stored .docx/.xlsx file.</summary>
public class DocumentPreviewPageViewModel
{
	public int UploadId { get; set; }

	public int DocId { get; set; }

	public string FileName { get; set; } = string.Empty;

	public ExtractedDocumentContent? Content { get; set; }

	public string? ErrorMessage { get; set; }
}