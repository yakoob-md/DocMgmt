
using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Http;

namespace DocMgmt.Models;

public class DocumentUploadRequest
{
	[Range(1, int.MaxValue, ErrorMessage = "Enter a valid document ID.")]
	public int DocId { get; set; }

	[Required(ErrorMessage = "Please select a file.")]
	public IFormFile? File { get; set; }
}
