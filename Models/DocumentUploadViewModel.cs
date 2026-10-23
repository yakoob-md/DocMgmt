
using Microsoft.AspNetCore.Http;

namespace DocMgmt.Models;

public class DocumentUploadViewModel
{
	public int DocId { get; set; }

	public IFormFile? File { get; set; }

	public DocumentUploadInfo? CurrentUpload { get; set; }

	public List<DocumentUploadInfo> UploadHistory { get; set; } = [];
}

public class DocumentUploadInfo
{
	public int UploadId { get; set; }

	public int DocId { get; set; }

	public string FileName { get; set; } = "";

	public string ContentType { get; set; } = "application/octet-stream";

	public long FileSizeBytes { get; set; }

	public DateTime? UploadedAt { get; set; }
}
