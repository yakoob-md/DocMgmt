using DocMgmt.Models;
using DocMgmt.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Xml;

namespace DocMgmt.Controllers
{
	public class DocumentsController : Controller
	{
		private readonly DocumentUploadService _uploadService;
		private readonly ILogger<DocumentsController> _logger;
		private readonly NativeDocumentParser _documentParser;

		public DocumentsController(
			DocumentUploadService uploadService,
			ILogger<DocumentsController> logger,
			NativeDocumentParser documentParser)
		{
			_uploadService = uploadService;
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
		[RequestSizeLimit(FileValidator.MaxFileBytes)]
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

		// --- PIPELINE A: PREVIEW BEFORE SUBMISSION (file is parsed in memory, nothing is stored) ---
		[HttpPost]
		[RequestSizeLimit(FileValidator.MaxFileBytes)]
		public async Task<IActionResult> PreviewBeforeSubmit(IFormFile? file, CancellationToken cancellationToken)
		{
			if (file is null || file.Length == 0)
				return PreviewError(StatusCodes.Status400BadRequest, "The selected file is empty.");

			if (file.Length > FileValidator.MaxFileBytes)
				return PreviewError(StatusCodes.Status413PayloadTooLarge, "The file is larger than the 100 MB limit.");

			var extension = Path.GetExtension(file.FileName).ToLowerInvariant();

			if (!FileValidator.IsOpenXml(extension))
				return PreviewError(StatusCodes.Status400BadRequest, "Server-side preview is only used for .docx and .xlsx files.");

			try
			{
				await using var memoryStream = new MemoryStream();
				await file.CopyToAsync(memoryStream, cancellationToken);
				memoryStream.Position = 0;

				if (!FileValidator.HasValidSignature(memoryStream, extension))
					return PreviewError(StatusCodes.Status422UnprocessableEntity,
						"This file is not a real Word/Excel file. Its contents do not match its extension.");

				var content = _documentParser.ParseFromStream(memoryStream, extension);
				return PartialView("_DocumentPreviewPartial", content);
			}
			catch (Exception ex) when (ex is InvalidDataException or XmlException)
			{
				_logger.LogWarning(ex, "Could not parse {FileName} for preview.", file.FileName);
				return PreviewError(StatusCodes.Status422UnprocessableEntity,
					"The file looks damaged or is not a valid Word/Excel file, so it cannot be previewed.");
			}
			catch (OperationCanceledException)
			{
				return new EmptyResult();
			}
			catch (Exception ex)
			{
				_logger.LogError(ex, "Unexpected preview failure for {FileName}.", file.FileName);
				return PreviewError(StatusCodes.Status500InternalServerError, "The preview could not be generated.");
			}
		}

		// --- PIPELINE B: PREVIEW AFTER SUBMISSION (reads the stored bytes from SQL Server) ---
		// embed=true  -> returns just the HTML fragment for the side panel
		// embed=false -> returns a full page (used by the "Tab" link)
		[HttpGet]
		public async Task<IActionResult> Preview(int uploadId, bool embed = false, CancellationToken cancellationToken = default)
		{
			if (uploadId <= 0) return BadRequest("A valid upload ID is required.");

			var upload = await _uploadService.GetUploadFileAsync(uploadId, cancellationToken);
			if (upload is null) return NotFound("The requested upload was not found.");

			var extension = Path.GetExtension(upload.FileName).ToLowerInvariant();

			if (FileValidator.IsOpenXml(extension))
			{
				var page = new DocumentPreviewPageViewModel
				{
					UploadId = upload.UploadId,
					DocId = upload.DocId,
					FileName = upload.FileName
				};

				try
				{
					using var memoryStream = new MemoryStream(upload.FileData, writable: false);
					page.Content = _documentParser.ParseFromStream(memoryStream, extension);
				}
				catch (Exception ex) when (ex is InvalidDataException or XmlException)
				{
					_logger.LogWarning(ex, "Stored upload {UploadId} could not be parsed.", uploadId);
					page.ErrorMessage = "The stored file looks damaged or is not a valid Word/Excel file, so it cannot be previewed. You can still download it.";
				}
				catch (Exception ex)
				{
					_logger.LogError(ex, "Unexpected preview failure for stored upload {UploadId}.", uploadId);
					page.ErrorMessage = "The preview could not be generated. You can still download the file.";
				}

				if (embed)
				{
					return page.Content is not null
						? PartialView("_DocumentPreviewPartial", page.Content)
						: PreviewError(StatusCodes.Status422UnprocessableEntity, page.ErrorMessage ?? "Preview failed.");
				}

				return View("PreviewPage", page);
			}

			if (!FileValidator.IsInlineSafe(extension))
				return StatusCode(StatusCodes.Status415UnsupportedMediaType,
					"Inline preview is not supported for this file type. Use Download instead.");

			var contentType = FileValidator.GetContentType(extension);
			if (contentType.StartsWith("text/", StringComparison.Ordinal))
				contentType += "; charset=utf-8";

			Response.Headers["Content-Disposition"] = "inline";
			Response.Headers["X-Content-Type-Options"] = "nosniff";

			return File(upload.FileData, contentType);
		}

		[HttpGet]
		public async Task<IActionResult> Download(int uploadId, CancellationToken cancellationToken)
		{
			if (uploadId <= 0) return BadRequest("A valid upload ID is required.");

			var upload = await _uploadService.GetUploadFileAsync(uploadId, cancellationToken);
			if (upload is null) return NotFound("The requested upload was not found.");

			var extension = Path.GetExtension(upload.FileName).ToLowerInvariant();
			Response.Headers["X-Content-Type-Options"] = "nosniff";

			return File(upload.FileData, FileValidator.GetContentType(extension), upload.FileName);
		}

		// Plain-text error body: the page script shows it to the user as-is.
		private ContentResult PreviewError(int statusCode, string message) =>
			new() { StatusCode = statusCode, Content = message, ContentType = "text/plain; charset=utf-8" };
	}
}