namespace DocMgmt.Models;

/// <summary>A stored upload including its bytes (used for preview and download).</summary>
public sealed class DocumentUploadFile
{
	public int UploadId { get; set; }

	public int DocId { get; set; }

	public string FileName { get; set; } = string.Empty;

	public string ContentType { get; set; } = "application/octet-stream";

	public byte[] FileData { get; set; } = Array.Empty<byte>();
}