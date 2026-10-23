
using Microsoft.Data.SqlClient;

namespace DocMgmt.Data;

public interface IDbConnectionFactory
{
	SqlConnection CreateConnection();
}
