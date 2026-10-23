
using DocMgmt.Data;
using Microsoft.AspNetCore.Mvc;
using Dapper;

namespace DocMgmt.Controllers;

[Route("database-test")]
public class DatabaseTestController : Controller
{
	private readonly IDbConnectionFactory _connectionFactory;

	public DatabaseTestController(
		IDbConnectionFactory connectionFactory)
	{
		_connectionFactory = connectionFactory;
	}

	[HttpGet]
	public async Task<IActionResult> Index()
	{
		try
		{
			using var connection =
				_connectionFactory.CreateConnection();

			await connection.OpenAsync();

			var databaseName = await connection.QuerySingleAsync<string>(
				"SELECT DB_NAME();");

			var tablesExist = await connection.QuerySingleAsync<bool>(
				"""
                SELECT CAST(
                    CASE WHEN
                        OBJECT_ID('dbo.Documents', 'U') IS NOT NULL
                        AND OBJECT_ID('dbo.DocumentUploads', 'U') IS NOT NULL
                    THEN 1 ELSE 0 END
                AS bit);
                """);

			return Content(
				$"SQL connection successful. Database: {databaseName}. " +
				$"Required tables exist: {tablesExist}");
		}
		catch (Exception ex)
		{
			// Log the exception in production rather than exposing
			// database details to users.
			return Problem(
				title: "Database connection test failed.",
				detail: ex.Message,
				statusCode: 500);
		}
	}
}
