
using Microsoft.Data.SqlClient;

namespace DocMgmt.Data;

public class SqlConnectionFactory : IDbConnectionFactory
{
	private readonly string _connectionString;

	public SqlConnectionFactory(IConfiguration configuration)
	{
		_connectionString =
			configuration.GetConnectionString("DefaultConnection")
			?? throw new InvalidOperationException(
				"Connection string 'DefaultConnection' was not found.");
	}

	public SqlConnection CreateConnection()
	{
		return new SqlConnection(_connectionString);
	}
}
