
Imports System.IO
Imports System.Data.OleDb
Module HouseHoldManagment_Module

    Public currentUser As String ' Global variable for logged-in user
    Public Property conn As New OleDbConnection(connectionString)
    ' Connection string using relative path to the database
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Mudzunga\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb;Persist Security Info=False;"

    ' Function to create and return a connection object
    Public Function GetConnection() As OleDbConnection
        Return New OleDbConnection(connectionString)
    End Function
    Public MealPlanId As Integer = 1

    Public Function MealPlanForm1() As String
        Dim formattedItemNumber As String = "ITM-" & MealPlanId.ToString("D3")
        MealPlanId += 1
        Return formattedItemNumber



    End Function
End Module
