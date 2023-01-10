using System.Data;
using System.Data.OleDb;

namespace ClassLibrary1
{
    public class Class1
    {
        public DataTable read(string connectionString)
        {

            OleDbConnection _connection = new OleDbConnection(connectionString);
            _connection.Open();

            OleDbDataAdapter adapter = new OleDbDataAdapter("Select * from [Sheet1$]", _connection);
            DataSet ds = new DataSet();

            adapter.Fill(ds);
            DataTable dataTable = ds.Tables[0];
            if (dataTable == null || dataTable.Columns.Count == 0)
            {
                throw new Exception("Blank file selected! Choose another file.");
            }
            return dataTable;
        }
    }
}