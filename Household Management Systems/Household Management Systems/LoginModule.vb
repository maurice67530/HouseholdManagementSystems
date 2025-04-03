Imports System.IO
Imports System.Data.OleDb

Module LoginModule

    Dim conn As New OleDbConnection(LoginModule.connectionString)
    Public Const connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Zwivhuya\Desktop\HouseholdManagementSystems\Household Management Systems\Household Management Systems\HMS.accdb"

End Module
