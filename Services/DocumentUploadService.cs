using Dapper;
using DocMgmt.Data;
using DocMgmt.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.Data.SqlClient;
using System.Data;

namespace DocMgmt.Services;

public class DocumentUploadService
{
	private readonly IDbConnectionFactory _connectionFactory;

	public DocumentUploadService(IDbConnectionFactory connectionFactory)
	{
		_connectionFactory = connectionFactory;
	}

	public async Task UploadAsync(
		int docId,
		IFormFile? file,
		CancellationToken cancellationToken = default)
	{
		if (docId <= 0)
			throw new ArgumentException("Invalid document ID.");

		if (file is null || file.Length == 0)
			throw new ArgumentException("Select a non-empty file.");

		if (file.Length > FileValidator.MaxFileBytes)
			throw new ArgumentException("The file is larger than the 100 MB limit.");

		// Never let a client-supplied path become the stored file name.
		string fileName = Path.GetFileName(file.FileName);

		if (string.IsNullOrWhiteSpace(fileName))
			throw new ArgumentException("Invalid filename.");

		if (fileName.Length > 255)
			throw new ArgumentException("The filename is too long.");

		var extension = Path.GetExtension(fileName).ToLowerInvariant();

		// The content type comes from the extension, not from the browser, so it can't be spoofed
		// and always fits the varchar(100) column.
		string contentType = FileValidator.GetContentType(extension);

		await using var fileStream = file.OpenReadStream();

		if (!FileValidator.HasValidSignature(fileStream, extension))
			throw new ArgumentException("File extension mismatch: the file contents do not match its extension.");

		if (FileValidator.IsOpenXml(extension) && !NativeDocumentParser.HasRequiredParts(fileStream, extension))
			throw new ArgumentException("The file is not a valid Word or Excel document.");

		fileStream.Position = 0; // validation moved the stream; the insert must start from byte 0

		await using var connection = _connectionFactory.CreateConnection();
		await connection.OpenAsync(cancellationToken);

		await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);

		try
		{
			int? documentId = await connection.QuerySingleOrDefaultAsync<int?>(
				new CommandDefinition(
					"""
					SELECT id
					FROM dbo.Documents
					WHERE id = @DocId;
					""",
					new { DocId = docId },
					transaction,
					cancellationToken: cancellationToken));

			if (documentId is null)
				throw new ArgumentException("The specified document was not found.");

			// ADO.NET streams the file into varbinary(max) without a second copy in memory.
			await using var command = new SqlCommand(
				"""
				INSERT INTO dbo.DocumentUploads
				(
					doc_id,
					original_file_name,
					file_content_type,
					file_size_bytes,
					file_data,
					uploaded_at
				)
				VALUES
				(
					@DocId,
					@FileName,
					@ContentType,
					@FileSize,
					@FileData,
					SYSUTCDATETIME()
				);
				""",
				connection,
				transaction);

			command.Parameters.Add("@DocId", SqlDbType.Int).Value = docId;
			command.Parameters.Add("@FileName", SqlDbType.NVarChar, 255).Value = fileName;
			command.Parameters.Add("@ContentType", SqlDbType.VarChar, 100).Value = contentType;
			command.Parameters.Add("@FileSize", SqlDbType.BigInt).Value = file.Length;
			command.Parameters.Add("@FileData", SqlDbType.VarBinary, -1).Value = fileStream;

			await command.ExecuteNonQueryAsync(cancellationToken);

			int rowsUpdated = await connection.ExecuteAsync(
				new CommandDefinition(
					"""
					UPDATE dbo.Documents
					SET status = @Status
					WHERE id = @DocId;
					""",
					new { DocId = docId, Status = "Submitted" },
					transaction,
					cancellationToken: cancellationToken));

			if (rowsUpdated != 1)
				throw new InvalidOperationException("The document status could not be updated.");

			await transaction.CommitAsync(cancellationToken);
		}
		catch
		{
			await transaction.RollbackAsync(CancellationToken.None);
			throw;
		}
	}

	public async Task<bool> DocumentExistsAsync(
		int docId,
		CancellationToken cancellationToken = default)
	{
		await using var connection = _connectionFactory.CreateConnection();
		await connection.OpenAsync(cancellationToken);

		return await connection.QuerySingleAsync<bool>(
			new CommandDefinition(
				"""
				SELECT CASE WHEN EXISTS (SELECT 1 FROM dbo.Documents WHERE id = @DocId)
				THEN CAST(1 AS bit) ELSE CAST(0 AS bit) END;
				""",
				new { DocId = docId },
				cancellationToken: cancellationToken));
	}

	public async Task<List<DocumentUploadInfo>> GetUploadHistoryAsync(
		int docId,
		CancellationToken cancellationToken = default)
	{
		await using var connection = _connectionFactory.CreateConnection();
		await connection.OpenAsync(cancellationToken);

		var history = await connection.QueryAsync<DocumentUploadInfo>(
			new CommandDefinition(
				"""
				SELECT
					upload_id AS UploadId,
					doc_id AS DocId,
					original_file_name AS FileName,
					file_content_type AS ContentType,
					file_size_bytes AS FileSizeBytes,
					uploaded_at AS UploadedAt
				FROM dbo.DocumentUploads
				WHERE doc_id = @DocId
				ORDER BY upload_id DESC;
				""",
				new { DocId = docId },
				cancellationToken: cancellationToken));

		return history.ToList();
	}

	/// <summary>Loads one stored upload including its bytes, or null if it does not exist.</summary>
	public async Task<DocumentUploadFile?> GetUploadFileAsync(
		int uploadId,
		CancellationToken cancellationToken = default)
	{
		await using var connection = _connectionFactory.CreateConnection();
		await connection.OpenAsync(cancellationToken);

		return await connection.QuerySingleOrDefaultAsync<DocumentUploadFile>(
			new CommandDefinition(
				"""
				SELECT
					upload_id AS UploadId,
					doc_id AS DocId,
					original_file_name AS FileName,
					file_content_type AS ContentType,
					file_data AS FileData
				FROM dbo.DocumentUploads
				WHERE upload_id = @UploadId;
				""",
				new { UploadId = uploadId },
				cancellationToken: cancellationToken));
	}
}