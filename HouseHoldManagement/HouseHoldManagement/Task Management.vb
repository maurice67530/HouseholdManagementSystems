Imports System.Data.OleDb

Public Class Task_Management
    Private Sub Task_Management_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim tooltip As New ToolTip
        tooltip.SetToolTip(Button1, "Submit")
        tooltip.SetToolTip(Button4, "Refresh")
        tooltip.SetToolTip(Button3, "Delete")
        tooltip.SetToolTip(Button2, "Edit")
        tooltip.SetToolTip(Button8, "Dashboard")
        tooltip.SetToolTip(Button5, "Filter")
        tooltip.SetToolTip(Button6, "Sort")

        ComboBox1.Items.AddRange(New String() {"Low", "Medium", "High"})
        ComboBox2.Items.AddRange(New String() {"Not started", "In progress", "Completed"})
    End Sub
    Private Sub ComboBox3_click(sender As Object, e As EventArgs) Handles ComboBox3.Click
        PopulateComboboxFromDatabase(ComboBox3)
    End Sub
    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
        Try
            Debug.WriteLine("populating combobox from database successfully!")
            ' 1. Open the database connection  
            conn.Open()

            ' 2. Retrieve the FirstName and LastName columns from the PersonalDetails table  
            Dim query As String = "SELECT FirstName, LastName FROM Person"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            ' 3. Bind the retrieved data to the combobox  
            comboBox.Items.Clear()
            While reader.Read()
                comboBox.Items.Add($"{reader("FirstName")} {reader("LastName")}")
            End While

            ' 4. Close the database connection  
            reader.Close()

        Catch ex As Exception
            Debug.WriteLine($"form loaded unsuccessful")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("An unexpected error occurred while populating the data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Close the database connection  
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class