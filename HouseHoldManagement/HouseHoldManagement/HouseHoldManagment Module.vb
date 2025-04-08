
Imports System.IO
Imports System.Data.OleDb
Module HouseHoldManagment_Module

    Public currentUser As String ' Global variable for logged-in user

    ' Connection string using relative path to the database
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Mudzunga\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb;Persist Security Info=False;"

    ' Function to create and return a connection object
    Public Function GetConnection() As OleDbConnection
        Return New OleDbConnection(connectionString)
    End Function

End Module
