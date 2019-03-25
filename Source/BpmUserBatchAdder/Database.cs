using System.Data;
using System.Data.SqlClient;

namespace BpmUserBatchAdder {
    public static class Database {
        public static DataTable GetDataTable(SqlCommand command, string sqlStatement) {
            command.CommandText = sqlStatement;
            var reader = command.ExecuteReader();
            var dataTable = new DataTable();
            dataTable.Load(reader);
            reader.Dispose();

            return dataTable;
        }

        public static bool RunSql(SqlCommand command, string sqlStatement) {
            command.CommandText = sqlStatement;

            return command.ExecuteNonQuery() != 0 ? true : false;
        }
    }    
}