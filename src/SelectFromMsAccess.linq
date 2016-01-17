<Query Kind="Statements">
  <Namespace>System.Data.Odbc</Namespace>
  <Namespace>System.Data.OleDb</Namespace>
</Query>

var db = @"C:\Sandpit\TestDb.accdb";
var sql = "SELECT * FROM tblTest";

// http://www.microsoft.com/en-us/download/confirmation.aspx?id=23734
var connString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};", db);

using(var conn = new OleDbConnection(connString)) 
{
	var da = new OleDbDataAdapter(sql, conn);
	var ds = new DataSet();

	da.Fill(ds, "tbl");
	ds.Dump();
}