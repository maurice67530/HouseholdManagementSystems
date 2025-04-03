
Imports System.Data.OleDb

Module household
    Public Property conn As New OleDbConnection(connectionString)
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Murangi\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
End Module
