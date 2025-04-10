
Imports System.IO
Imports System.Data.OleDb
Module HouseHoldManagment_Module

    Public currentUser As String ' Global variable for logged-in user
    Public Property conn As New OleDbConnection(connectionString)
    ' Connection string using relative path to the database
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Zwivhuya\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb;Persist Security Info=False;"

    ' Function to create and return a connection object
    Public Function GetConnection() As OleDbConnection
        Return New OleDbConnection(connectionString)
    End Function

End Module
Module InventoryModule
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Dongola\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"

    Public Sub FilterInventory(Category As String, Unit As String)
        Dim taskTable As New DataTable
        Dim conn As New OleDbConnection(connectionString)
        Try
            conn.Open()
            Dim query As String = "SELECT * FROM Inventory WHERE 1=1"

            If Not String.IsNullOrEmpty(Category) Then
                query &= " AND Category = @Category"
            End If

            If Not String.IsNullOrEmpty(Unit) Then
                query &= " AND Unit = @Unit"
            End If

            Dim command As New OleDb.OleDbCommand(query, conn)

            If Not String.IsNullOrEmpty(Category) Then
                command.Parameters.AddWithValue("@Category", Category)
            End If

            If Not String.IsNullOrEmpty(Unit) Then
                command.Parameters.AddWithValue("@Unit", Unit)
            End If

            Dim adapter As New OleDb.OleDbDataAdapter(command)
            adapter.Fill(taskTable)
            Inventory.DataGridView1.DataSource = taskTable
        Catch ex As Exception
            MsgBox("Error filtering Inventory: " & ex.Message, MsgBoxStyle.Critical, "Database Error")

        Finally
            conn.Close()
        End Try
    End Sub
End Module
