namespace DocMgmt.Models
{
    public class Document
    {
		public int Id { get; set; }

		public int EmpId { get; set; }

		public string Category { get; set; } = string.Empty;

		public DateTime StartDate { get; set; }

		public string CategoryId { get; set; } = string.Empty;

		public int ReqId { get; set; }

		public string ReqName { get; set; } = string.Empty;

		public string Status { get; set; } = string.Empty;

		public string? Remarks { get; set; }

		// Related uploaded files.
		public List<DocumentUpload> DocumentUploads { get; set; } = new();
	}
}
