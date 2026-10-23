
using Dapper;
using DocMgmt.Data;
using DocMgmt.Data;
using DocMgmt.Models;

namespace DocMgmt.Repositories;

public class DocumentRepository
{
	private readonly IDbConnectionFactory _connectionFactory;

	public DocumentRepository(IDbConnectionFactory connectionFactory)
	{
		_connectionFactory = connectionFactory;
	}

	public async Task<Document?> GetByIdAsync(int id)
	{
		const string sql = """
            SELECT
                id AS Id,
                emp_id AS EmpId,
                category AS Category,
                start_date AS StartDate,
                category_id AS CategoryId,
                req_id AS ReqId,
                req_name AS ReqName,
                status AS Status,
                remarks AS Remarks
            FROM dbo.Documents
            WHERE id = @Id;
            """;

		using var connection = _connectionFactory.CreateConnection();

		return await connection.QuerySingleOrDefaultAsync<Document>(
			sql,
			new { Id = id });
	}

	public async Task<IEnumerable<Document>> GetAllAsync()
	{
		const string sql = """
            SELECT
                id AS Id,
                emp_id AS EmpId,
                category AS Category,
                start_date AS StartDate,
                category_id AS CategoryId,
                req_id AS ReqId,
                req_name AS ReqName,
                status AS Status,
                remarks AS Remarks
            FROM dbo.Documents
            ORDER BY id;
            """;

		using var connection = _connectionFactory.CreateConnection();

		return await connection.QueryAsync<Document>(sql);
	}
}
