namespace DocMgmt.Models
{
    public class DocumentUpload
    {
		public int UploadId { get; set; }

		public int DocId { get; set; }

		public string OriginalFileName { get; set; } = string.Empty;

		public string FileContentType { get; set; } = string.Empty;

		public long FileSizeBytes { get; set; }

		public byte[] FileData { get; set; } = Array.Empty<byte>();

		public DateTime? UploadedAt { get; set; }

		// Parent document, if loaded by the application.
		public Document? Document { get; set; }
	}
}
