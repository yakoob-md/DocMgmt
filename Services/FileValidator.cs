namespace DocMgmt.Services;

/// <summary>Shared, dependency-free file checks used by both the upload and preview paths.</summary>
public static class FileValidator
{
	public const long MaxFileBytes = 100L * 1024 * 1024; // 100 MB

	private static readonly Dictionary<string, string> ContentTypes = new(StringComparer.OrdinalIgnoreCase)
	{
		[".docx"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
		[".xlsx"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
		[".csv"] = "text/csv"
	};

	// Types that are safe to stream back inline (no scripts can run from them).
	private static readonly HashSet<string> InlineSafe = new(StringComparer.OrdinalIgnoreCase)
	{
		".pdf", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".mp3", ".wav", ".mp4", ".webm", ".txt", ".csv"
	};

	public static string GetContentType(string extension) =>
		ContentTypes.TryGetValue(extension, out var ct) ? ct : "application/octet-stream";

	public static bool IsOpenXml(string extension) =>
		extension.Equals(".docx", StringComparison.OrdinalIgnoreCase) ||
		extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase);

	public static bool IsInlineSafe(string extension) => InlineSafe.Contains(extension);

	/// <summary>Checks the leading "magic" bytes. Extensions without a known signature pass.</summary>
	public static bool HasValidSignature(Stream stream, string extension)
	{
		byte[]? expected = extension.ToLowerInvariant() switch
		{
			".docx" or ".xlsx" => new byte[] { 0x50, 0x4B, 0x03, 0x04 },
			".pdf" => new byte[] { 0x25, 0x50, 0x44, 0x46 },
			".png" => new byte[] { 0x89, 0x50, 0x4E, 0x47 },
			".jpg" or ".jpeg" => new byte[] { 0xFF, 0xD8, 0xFF },
			".gif" => new byte[] { 0x47, 0x49, 0x46, 0x38 },
			_ => null
		};

		if (expected is null) return true;

		long original = stream.Position;
		try
		{
			stream.Position = 0;
			var actual = new byte[expected.Length];
			int read = stream.Read(actual, 0, actual.Length);
			return read == expected.Length && actual.AsSpan().SequenceEqual(expected);
		}
		finally
		{
			stream.Position = original;
		}
	}
}