Imports System.Data.OleDb
Imports System.Windows.Forms
Imports System.Security.Cryptography
Imports System.Text

Module Module1
    ' Connection string using relative path to the database
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Raphalalani\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"

    ' Function to create and return a connection object
    Public Function GetConnection() As OleDbConnection
        Return New OleDbConnection(connectionString)
    End Function
End Module
