using Dapper;
using DocMgmt.Data;
using DocMgmt.Models;
using DocMgmt.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace DocMgmt.Controllers
{
	public class DocumentsController : Controller
	{
		private readonly DocumentUploadService _uploadService;
		private readonly IDbConnectionFactory _connectionFactory;
		private readonly ILogger<DocumentsController> _logger;
		private readonly NativeDocumentParser _documentParser;

		public DocumentsController(
			DocumentUploadService uploadService,
			IDbConnectionFactory connectionFactory,
			ILogger<DocumentsController> logger,
			NativeDocumentParser documentParser)
		{
			_uploadService = uploadService;
			_connectionFactory = connectionFactory;
			_logger = logger;
			_documentParser = documentParser;
		}

		[HttpGet]
		public async Task<IActionResult> Upload(int? docId, CancellationToken cancellationToken)
		{
			var model = new DocumentUploadViewModel
			{
				DocId = docId ?? 0
			};

			if (docId.HasValue && docId.Value > 0)
			{
				if (!await _uploadService.DocumentExistsAsync(docId.Value, cancellationToken))
				{
					return NotFound("The requested document was not found.");
				}

				await LoadHistoryAsync(model, cancellationToken);
			}

			return View(model);
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public async Task<IActionResult> Upload(DocumentUploadViewModel model, CancellationToken cancellationToken)
		{
			if (model.File is null || model.File.Length == 0)
			{
				ModelState.AddModelError(nameof(model.File), "Select a file before uploading a new version.");
			}

			if (model.DocId <= 0)
			{
				ModelState.AddModelError(nameof(model.DocId), "A valid document ID is required.");
			}

			if (ModelState.IsValid)
			{
				try
				{
					await _uploadService.UploadAsync(model.DocId, model.File, cancellationToken);

					TempData["Success"] = "The new version was uploaded successfully. Previous versions have been retained.";
					return RedirectToAction(nameof(Upload), new { docId = model.DocId });
				}
				catch (ArgumentException ex)
				{
					ModelState.AddModelError(string.Empty, ex.Message);
				}
				catch (Exception ex)
				{
					_logger.LogError(ex, "Upload failed for document {DocId}.", model.DocId);
					ModelState.AddModelError(string.Empty, "The upload failed. Please try again.");
				}
			}

			if (model.DocId > 0 && await _uploadService.DocumentExistsAsync(model.DocId, cancellationToken))
			{
				await LoadHistoryAsync(model, cancellationToken);
			}

			return View(model);
		}

		private async Task LoadHistoryAsync(DocumentUploadViewModel model, CancellationToken cancellationToken)
		{
			model.UploadHistory = await _uploadService.GetUploadHistoryAsync(model.DocId, cancellationToken);
			model.CurrentUpload = model.UploadHistory.FirstOrDefault();
		}

		// --- PIPELINE A: BEFORE-SUBMISSION PREVIEW (With Verification Checking) ---
		[HttpPost]
		[RequestSizeLimit(104857600)]
		public IActionResult PreviewBeforeSubmit(IFormFile file)
		{
			if (file == null || file.Length == 0)
			{
				return BadRequest("Invalid or empty file data.");
			}

			var extension = Path.GetExtension(file.FileName).ToLowerInvariant();

			if (extension == ".docx" || extension == ".xlsx")
			{
				try
				{
					using var memoryStream = new MemoryStream();
					file.CopyTo(memoryStream);
					memoryStream.Position = 0;

					// Verify signature matching before extracting structure
					byte[] expectedSig = new byte[] { 0x50, 0x4B, 0x03, 0x04 };
					byte[] actualSig = new byte[4];
					memoryStream.Read(actualSig, 0, 4);
					memoryStream.Position = 0; // Reset after reading signature

					if (!actualSig.SequenceEqual(expectedSig))
					{
						return BadRequest("Security Block: File manipulation detected. The binary header does not match a valid OpenXML Zip structure.");
					}

					var parsedContent = _documentParser.ParseFromStream(memoryStream, extension);
					return PartialView("_DocumentPreviewPartial", parsedContent);
				}
				catch (Exception ex)
				{
					return BadRequest($"Parsing Failure: {ex.Message}");
				}
			}

			return BadRequest("Standard media formats are viewed directly client-side.");
		}

		// --- PIPELINE B: AFTER-SUBMISSION PREVIEW (Serves raw bytes straight to UI layout) ---
		[HttpGet]
		public async Task<IActionResult> Preview(int uploadId, CancellationToken cancellationToken)
		{
			if (uploadId <= 0) return BadRequest();

			await using var connection = _connectionFactory.CreateConnection();
			await connection.OpenAsync(cancellationToken);

			var upload = await connection.QuerySingleOrDefaultAsync<PreviewFile>(
				new CommandDefinition(
					"""
                    SELECT
                        upload_id AS UploadId,
                        doc_id AS DocId,
                        original_file_name AS FileName,
                        file_data AS FileData
                    FROM dbo.DocumentUploads
                    WHERE upload_id = @UploadId;
                    """,
					new { UploadId = uploadId },
					cancellationToken: cancellationToken));

			if (upload is null) return NotFound();

			var extension = Path.GetExtension(upload.FileName).ToLowerInvariant();

			if (extension == ".docx" || extension == ".xlsx")
			{
				using var memoryStream = new MemoryStream(upload.FileData);
				var parsedContent = _documentParser.ParseFromStream(memoryStream, extension);
				return PartialView("_DocumentPreviewPartial", parsedContent);
			}

			var types = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
			{
				[".pdf"] = "application/pdf",
				[".png"] = "image/png",
				[".jpg"] = "image/jpeg",
				[".jpeg"] = "image/jpeg",
				[".gif"] = "image/gif",
				[".webp"] = "image/webp",
				[".mp3"] = "audio/mpeg",
				[".wav"] = "audio/wav",
				[".mp4"] = "video/mp4",
				[".webm"] = "video/webm",
				[".txt"] = "text/plain",
				[".csv"] = "text/plain"
			};

			if (!types.TryGetValue(extension, out var contentType))
			{
				return StatusCode(StatusCodes.Status415UnsupportedMediaType, "Inline preview is not supported for this file type.");
			}

			Response.Headers["Content-Disposition"] = "inline";
			Response.Headers["X-Content-Type-Options"] = "nosniff";

			return File(upload.FileData, contentType);
		}

		private sealed class PreviewFile
		{
			public int UploadId { get; set; }
			public int DocId { get; set; }
			public string FileName { get; set; } = "";
			public byte[] FileData { get; set; } = [];
		}
	}
}
