using Dapper;
using DocMgmt.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

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
		// Validate the document ID.
		if (docId <= 0)
			throw new ArgumentException("Invalid document ID.");

		// Validate that a file was selected.
		if (file is null || file.Length == 0)
			throw new ArgumentException("Select a non-empty file.");

		// Do not allow a client-supplied path to become the filename.
		string fileName = Path.GetFileName(file.FileName);

		if (string.IsNullOrWhiteSpace(fileName))
			throw new ArgumentException("Invalid filename.");

		if (fileName.Length > 255)
			throw new ArgumentException("The filename is too long.");

		// --- SECURITY LAYER 1: Dynamic Content-Type Resolution ---
		string contentType = file.ContentType;
		var extension = Path.GetExtension(fileName).ToLowerInvariant();

		if (string.IsNullOrWhiteSpace(contentType) || contentType == "application/octet-stream")
		{
			contentType = extension switch
			{
				".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
				".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
				".pdf" => "application/pdf",
				".txt" => "text/plain",
				_ => "application/octet-stream"
			};
		}

		// Open the uploaded file stream.
		await using var fileStream = file.OpenReadStream();

		// --- SECURITY LAYER 2: First-Party Binary Signature Verification ---
		if (!VerifyFileSignatureNatively(fileStream, extension))
		{
			throw new ArgumentException("Security Warning: File extension mismatch. The internal file contents do not match its file extension.");
		}

		await using var connection = _connectionFactory.CreateConnection();
		await connection.OpenAsync(cancellationToken);

		using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);

		try
		{
			// Verify that the parent document exists.
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

			// Use ADO.NET to stream the file into varbinary(max).
			using var command = new SqlCommand(
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
			command.Parameters.Add("@FileName", SqlDbType.VarChar, 255).Value = fileName;
			command.Parameters.Add("@ContentType", SqlDbType.VarChar, 100).Value = contentType;
			command.Parameters.Add("@FileSize", SqlDbType.BigInt).Value = file.Length;

			SqlParameter fileParameter = command.Parameters.Add("@FileData", SqlDbType.VarBinary, -1);
			fileParameter.Value = fileStream;

			await command.ExecuteNonQueryAsync(cancellationToken);

			// Update the checklist status in the same transaction.
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

		var exists = await connection.QuerySingleAsync<bool>(
			new CommandDefinition(
				"""
            SELECT CASE WHEN EXISTS
            (
                SELECT 1
                FROM dbo.Documents
                WHERE id = @DocId
            )
            THEN CAST(1 AS bit)
            ELSE CAST(0 AS bit)
            END;
            """,
				new { DocId = docId },
				cancellationToken: cancellationToken));

		return exists;
	}

	public async Task<List<DocMgmt.Models.DocumentUploadInfo>> GetUploadHistoryAsync(
		int docId,
		CancellationToken cancellationToken = default)
	{
		await using var connection = _connectionFactory.CreateConnection();
		await connection.OpenAsync(cancellationToken);

		var history = await connection.QueryAsync<DocMgmt.Models.DocumentUploadInfo>(
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

	// First-Party Magic Number Check to verify authentic formats without software installations
	private bool VerifyFileSignatureNatively(Stream fileStream, string extension)
	{
		var validSignatures = new Dictionary<string, byte[]>
		{
			{ ".docx", new byte[] { 0x50, 0x4B, 0x03, 0x04 } }, // Zip format archive signature (PK..)
			{ ".xlsx", new byte[] { 0x50, 0x4B, 0x03, 0x04 } },
			{ ".pdf",  new byte[] { 0x25, 0x50, 0x44, 0x46 } }  // Native PDF header (%PDF)
		};

		string ext = extension.ToLowerInvariant();
		if (!validSignatures.ContainsKey(ext)) return true; // Let formats without simple headers pass (.txt, .csv)

		byte[] expectedSignature = validSignatures[ext];
		byte[] actualSignature = new byte[expectedSignature.Length];

		long originalPosition = fileStream.Position;
		fileStream.Position = 0; // Read from the absolute beginning

		int bytesRead = fileStream.Read(actualSignature, 0, actualSignature.Length);
		fileStream.Position = originalPosition; // Always reset stream marker position back immediately

		if (bytesRead < expectedSignature.Length) return false;

		return actualSignature.SequenceEqual(expectedSignature);
	}
}
